MOMA Accessions
==================================

Utility for migrating accession spreadsheets to ArchivesSpace.

To use this, you will need to first install NodeJS and NPM, and then install the "ArchiveSpace Command Line Interface", an experimental tool for working with the ArchivesSpace API using the command line and Javascript.

Usage:

    $ npm install as-cli
    $ git clone git@github.com:lcdhoffman/moma-accessions.git
    $ cd moma-accessions
    $ npm init
    $ npm install minimist
    $ as-cli run-script migrate.js --locations path/to/data.xlsx --collections path/to/more/data.xlsx




