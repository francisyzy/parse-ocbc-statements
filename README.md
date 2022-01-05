# Parse OCBC Statements

Well OCBC sends monthly statements in PDF. We `xlsx` or `csv` etc to do own analysis on own financial situation!

This parses the entire statement into one big `xlsx` sheet where you can filter/sort add your own stuff

See [sample](./sample) for a sample string of what `pdf-parse` package exports.

Afterwards, I just broke the string line by line and do some regex to find the relevant lines to add into `xlsx`

You can modify the script however you like. Can just remove the `everything.push(...statement);` and uncomment some lines to generate monthly statement by sheet if you like.

Free code just steal