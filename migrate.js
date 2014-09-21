#!/usr/bin/env node

var moment = require('moment');
var XLSX = require('xlsx');

var argv = require('minimist')(process.argv.slice(2));
var fs = require('fs');
var path = require('path');

Array.prototype.each = function(callback) {
  for (var i=0; i < this.length; i++) { callback(this[i]); } 
}


function AccessionConverter(instances, api) {
  var instances = instances;
  var api = api;

  var records = [];
  var classification_map = {};
  var headers = {};


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
            console.log("Saved Classification " + json.uri)
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

        console.log("Saved Location: " + json.uri);

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

      
      if (json.uri) {
        console.log("Saved Accession: " + json.uri);
      } else {
        console.log("Failed to save Accession");
        console.dir(json);
      }
    });
  };


  this.send = function() {

    records.each(function(rec) {

      var my_instances = instances[rec.title];

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
    E: function(obj, val) {
      obj.disposition = "RG Number " + val;
    },
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
      var status = "new";
      if (val === "No") {
        status = "completed";
      }

      obj.collection_management = {
        processing_status: status
      }
    },
    N: "access_restrictions_note",
    V: function(obj, val) {
      if (obj.general_note.length > 0) {
        obj.general_note = "\n\n" + obj.general_note;
      }
      obj.general_note = "Note:\n" + val + obj.general_note;
    },
    W: function(obj, val) {
      var ext = {
        "title": "Finding Aid Link",
        "location": val.replace(/^#/, '').replace(/#$/, '')
      }

      obj.external_documents.push(ext);

    }

  }

  this.setHeader = function (code, value) {
    var value = value.trim().replace(/^"/, '').replace(/"$/, '');
    headers[code] = value;
  }

  this.readCell = function (code, value) {
    var value = value.trim().replace(/^"/, '').replace(/"$/, '');

    if (value.length < 1) return;

    if (code == 'A') {
      records.push({
        general_note:"LEGACY EXCEL DATA:\n",
        instances: [],
        external_documents: []
      });
    };

    records[records.length - 1].general_note += (headers[code] + ": " + value + "\n");
    
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


module.exports = function(api) {

  var locations = XLSX.readFile(argv['locations']);
  var collections = XLSX.readFile(argv['collections']);
  var instances = {};

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

  var converter = new AccessionConverter(instances, api);

  collections.SheetNames.forEach(function(y) {
    var worksheet = collections.Sheets[y];

    for (z in worksheet) {
      if(z[0] === "!") continue;

      var v = JSON.stringify(worksheet[z].v);

      if(z[1] === "1" && z.length === 2) {
        converter.setHeader(z[0], v);
        } else {
          converter.readCell(z[0], v)
        }
    };
  });

  converter.send();
}

