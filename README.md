# timesheets-parser
HRStop Time Sheets CSV to EXCEL 

**Setup**
* Install GIT, NPM and Node
  * https://docs.npmjs.com/downloading-and-installing-node-js-and-npm
* Clone the repo
* run `npm install`

**Run the Script**
* run `node index.js -h` to see the options

Download the time sheets CSV from HRStop and give that file as input using `-i` option
<br/>

Example: `node index.js -i timesheets.csv`
<br/>

You can also specify output file name using `-o` option
<br/>

Example: `node index.js -i timesheets.csv -o Set2023TimeSheets.xlsx`
<br/>

Multiple entries for a day are combined by default, specify `-dc` to not combine
<br/>

Example: `node index.js -i timesheets.csv -dc`

<br/>
<br/>
<br/>
<b>Author</b>
N V Chalapathi Raju