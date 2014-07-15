var App = require('as_cli')['App'];
var Task = require('as_cli')['Task'];
var api = require('as_cli')['api'];
var logger = require('as_cli')['logger'];
var moment = require('moment');

var app = new App();
var XLSX = require('xlsx');

function AccessionConverter(instances) {
  var instances = instances;
  var records = [];
  var classification_map = {};


  var sendClassification = function (rec, callback) {
    if (typeof(rec.classification) === "string") {
      if (classification_map[rec.classification] ) {
        rec.classification = {ref: classification_map[rec.classification]};
          callback(rec);
        } else {
          api.createClassification({
            identifier: Object.keys(classification_map).length + "",
            title: rec.classification
          }, function(err, json) {
            if (err) throw "err " + err;
            rec.classification = {ref: json.uri}
            callback(rec);
          });
        }
      } else {
        callback(rec);
      }
  };


  var sendLocations = function (rec, instances, callback) {
    var instances = instances || [];
    var count = instances.length;

    if (count == 0) {
      callback(rec);
    }

    instances.each(function(i) {
      var loc = i.container.container_locations[0]._resolved;

      api.createLocation(loc, function(err, json) {
        if (err) throw "err " + err;

        i.container.container_locations[0].ref = json.uri;
        delete i.container.container_locations[0]._resolved;

        rec.instances.push(i);
        count--;

        if (count == 0) callback(rec);
      });
    });
  };


  var sendAccession = function (rec) {
    api.createAccession(rec, function(err, json) {
      if (err) throw "err " + err;

      logger.info("Saved Accession: " + json.uri);
    });
  };


  this.send = function() {

    logger.debug("Sending");
    logger.debug(classification_map);
    records.each(function(rec) {
      logger.debug(rec);

      var my_instances = instances[rec.title];
      logger.debug(my_instances);

      sendLocations(rec, my_instances, function(rec) {

        sendClassification(rec, function(rec) {

          sendAccession(rec);
        });
      });
    });
  };

    
  var map = {
    A: 'title',
    B: 'provenance',
    C: function(obj, val) {
      var date = moment(val);
      if (date.isValid()) {
        obj.accession_date = date.format("YYYY-MM-DD");
      }
    },
    D: function(obj, val) {
      var ids = val.split('.');
      i = 0;
      while (ids.length > 0) {
        obj["id_" + i] = ids.shift();
        i++;
      }
    },
    E: 'id_3',
    G: function(obj, val) {
      if (val === "TRUE") {
        obj.publish = true;
      }
    },
    H: function(obj, val) {
      if (val.match(/^\d{4}$/)) {
        obj.dates = [{
          date_type: "single",
          label: "other",
          begin: val
        }];
      }
    },
    I: function(obj, val) {
      if (val.match(/^\d{4}$/)) {
        if (obj.dates.length === 1) {
          obj.dates[0].date_type = "inclusive";
          obj.dates[0].end = val;
        }
      }
    },
    J: "classification",
    K: function(obj, val) {
      obj.extents = [{
        extent_type: "linear_feet",
        portion: "whole",
        number: val
      }]
    },
    M: function(obj, val) {
      logger.debug(val);
      var status = "new";
      if (val === "No") {
        logger.debug("-----");
        status = "completed";
      }

      obj.collection_management = {
        processing_status: status
      }
    }
  }

  this.readCell = function (code, value) {
    var value = value.trim().replace(/^"/, '').replace(/"$/, '');

    if (value.length < 1) return;

    if (code == 'A') {
      records.push({
        general_note:"",
        instances: []
      });
    };

    records[records.length - 1].general_note += ("Cell " + code + ": " + value + "\n");
    
    if (map[code]) {

      if (typeof(map[code]) == 'string') {
        records[records.length - 1][map[code]] = value;
      } else {
        map[code](records[records.length - 1], value);
      }
    }
  };

  this.inspect = function (log) {
    records.each(function(r) {
      log.debug(r);
    });
  };

};



app.extend(new Task('moma-accessions', function(argv) {
  var locations = XLSX.readFile(argv['locations']);
  var collections = XLSX.readFile(argv['collections']);
  var instances = {}

  var loc;
  var cont_loc;
  var cont;

  locations.SheetNames.forEach(function(y) {
    var worksheet = locations.Sheets[y];
    for (z in worksheet) {
      if(z[0] === "!") continue;
      if(z[1] === "1" && z.length === 2) continue;

      var v = JSON.stringify(worksheet[z].v);
      v = v.trim().replace(/^"/, '').replace(/"$/, '');

      switch(z[0]) {
      case 'A':

        loc = {
          external_ids: [{
            external_id: v,
            source: "Locations.xlsx"
          }]
        };

        break;

      case 'B':

        loc.building = v;
        loc.classification = "Building";

        // api.createLocation(loc, function(err, json) {          
        //   if (err) {
        //     console.log("Error " + err);
        //   } else {
        //     loc.uri = json.uri;
        //   }
        // });

        break;

      case 'C':
        cont_loc = {
          status: "current",
          start_date: "1929-11-07"
        }
        cont = {}
        if (v) cont_loc.note = v;

        cont_loc._resolved = loc;
        cont.type_1 = "box";
        cont.indicator_1 = "1";
        cont.container_locations = [cont_loc];

        break;

      case 'D':
        instances[v] = instances[v] || [];
        instances[v].push({
          instance_type: "accession",
          container: cont
        });

        break;        
      }
    }
  });

  var converter = new AccessionConverter(instances);

  collections.SheetNames.forEach(function(y) {
    var worksheet = collections.Sheets[y];

    for (z in worksheet) {
      if(z[0] === "!") continue;
      if(z[1] === "1" && z.length === 2) continue;

//      if(z[1] === "3") break;

      var v = JSON.stringify(worksheet[z].v);

      logger.debug(z + "::" + v);

      converter.readCell(z[0], v)
    };
  });

  converter.inspect(logger);
  converter.send();

  // collections.SheetNames.forEach(function(y) {
  //   var worksheet = collections.Sheets[y];

  //   for (z in worksheet) {
  //     if(z[0] === "!") continue;
  //     if(z[1] === "1") continue;

  //     var v = JSON.stringify(worksheet[z].v);

  //     switch(z[0]) {
  //     case 'A':
  //       accession = {};
  //       accession.title = v;

  //       accession.instances = instances[v];

  //       break;

  //     case 'B':
  //       accession.provenance = v;
  //       break;

  //     case 'C':
  //       accession.general_note = v;
  //       break;

  //     case 'D':
  //       ids = v.split('.');
  //       i = 0;
  //       // while (ids.length > 0) {
  //       //   accession["id_" + i] = ids.shift();
  //       // }
        
        

  //     };

  //     logger.debug(accession);
  //   };
  // });


}).flags({
  locations: {
    required: true,
    boolean: false
  },
  collections: {
    required: true,
    boolean: false
  }
}));


app.run();

