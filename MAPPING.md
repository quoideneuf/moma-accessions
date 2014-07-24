Mappings for MOMA Accession Importer
===========================

# Locations.xlsx

| XLSX Cell | ASpace data location |
| --------- | ---------------------|
| A | `location.external_ids[0].external_id` |
| B | `location.building` |
| C | `accession.instances[0].container.container_locations[0].note` |
| D | Used to join the location to the accession if possible |


# Collections.xlsx

| XLSX Cell | ASpace data location |
| --------- | ---------------------|
| A | `accession.title` |
| B | `accession.provenance` |
| C | `accession.accession_date` |
| D | `accession.id_{0-2}` |
| E | `accession.id_3` |
| F | ignore |
| G | `accession.publish` |
| H | `accession.dates[0].begin` |
| I | `accession.dates[0].end` |
| J | `accession.classification` |
| K | `accession.extents[0].number` |
| L | ignore |
| M | `accession.collection_management.processing_status` |
| N | ignore |
| O | ignore |
| P | ignore |
| Q | ignore |
| R | ignore |
| S | ignore |
| T | ignore |
| U | ignore |
| V | `accession.general_note` |
| W | `accession.general_note` |

