# convert office files to mysql inserts becuase there are so many problems with converting to csv and importing its hilarious

# Expected File Format:
# tableName
# col1  col2  col3
# data  data  data
# data  data  data
# ...

# Output:
# INSERT INTO tableName (col1, col2, col3) VALUES (data, data, data);
# ...

# Usage:
# edit line 25 with file name of your target
# coffee convert.coffee > temp.sql
# mysql -u root -p isee_webapp < temp.sql

# @TODO make this more like a command line utility, accept file as input, have a help section, etc.

# DANGER!! <-- you'll need to debug this when you have internet
# node-xlsx failed to properly read a row that had duplicate numbers, e.g.
# questionStatsFake.ods line 136: 143  171 171 2 344
# gives column missing [ 'questionId', 'correct', 'incorrect', 'blank', 'total' ] [ 143, 171, 2, 344 ]
# what the shit!

xlsx = require "node-xlsx"

util = require "util"
fs = require "fs"

options = require("cli").execute {
  file: String,
  table: String,
  worksheet: [Number],
  columns: [Array],
  extras: [Object],
  h: [Boolean],
  help: [Boolean]
}

if options.h or options.help
  console.log """

Parses excel files and spits out mysql.  The updates you may
need to handwrite, but inserts are easy to generalize.  The
script is more stable if you are parsing csv files as there
are known bugs with the module for parsing xlsx files.
However it has the advantage of being able to deal with
multiple worksheets.

  file: String file name
  table: String name of table to insert or update
  worksheet: [Number] if xlsx file has worksheets, this is the index
  columns: [Array] indexes of desired columns, if not all
  extras: [Object] json of {col: data} for extra data to include


Usage:
  coffee convert.coffee --file SSAT_VR_Percentiles.xlsx --table \
    percentiles --worksheet 1 --insert --columns 1,2 1>temp.sql

"""
  process.exit(0)

console.log "/*\n#{process.argv.join(" ")}\n*/"

# these fields we are going to stuff into whatever is already in the csv
# we convert it to an array for easy concat
extraFields = []
extraData = []
if options.extras
  for key, value of options.extras
    extraFields.push key
    extraData.push value

# this is an array of fields that we want to keep, by index, and discard the rest
# we convert it to a map for easy access
indexMap = {}
if options.columns
  indexMap = options.columns.reduce (map, current)->
    map[current] = true
    return map
  , indexMap

# read either a csv file or an excel file and call the callback with the entire thing, split up by rows and cells
readFile = (file, callback)->
  fileType = file.split(".").pop()
  # read the csv file, this is stable and works
  if fileType is "csv"
    fs.readFile file, {encoding: 'utf8'}, (error, data)->
      if error
        console.error error
      else
        callback null, convert(data)

  # attempt to use xlsx, note the huge bug mentioned above...
  else
    obj = xlsx.parse(file)
    if options.worksheet isnt undefined and obj[options.worksheet]
      callback null, obj[options.worksheet].data
    else
      console.error "please specify a worksheet by Number"

# parse a csv text file into rows and columns:
# @param String
# @return array of rows with array of columns
convert = (data)->
  rows = data.split("\n").map (row)->
    output = []
    while row.length > 0
      if row[0] is "\""
        regexp = /[^\\]"(,|$)/
        adjust = 2
      else
        regexp = /(,|$)/
        adjust = 0

      index = row.search(regexp)
      if index is -1
        console.error "wtf regxp not found", regexp
      else
        elem = row.substring 0, (index+adjust)
        output.push elem
        row = row.substring(index + adjust + 1)
        #console.log "found", regexp, elem, row, row.length
    return output
  return rows

# options can specify to only include some of the columns from the dataset, by index
filterFields = (row, originalLength)->
  if options.columns
    return row.filter (elem, index)->
      # either it was marked as part of the filter, or its one of the extras
      return indexMap[index] is true or index >= originalLength
  else
    return row

# produce the query:
safeC = null
getInsert = (columns, row)->
  if not safeC
    safeC = columns.map((elem)-> return "#{elem}").join()
  safeR = row.map((elem)-> return "\"#{elem}\"").join()
  return "INSERT INTO #{options.table} (#{safeC}) VALUES (#{safeR});"

# options can stuff in extra static data that was not included in the file:
addExtras = (row, isColumns = false)->
  if options.extras
    if isColumns
      return row.concat extraFields
    else
      return row.concat extraData
  else
    return row

# do some custom junk, the user can override this to modify columns as they see fit:
custom = (row, isColumns = false)->
  return row

exports.execute = (custumFn = custom, queryFn = getInsert)->
  readFile options.file, (error, data)->
    colLength = 0
    colFixed = null

    for row, index in data
      # assume firt row is the column names:
      if index is 0
        columns = row.map (elem)->return "#{elem}".replace(/"/g, "").trim()
        colLength = columns.length
        colFixed = filterFields(custumFn(addExtras(columns, true), true), colLength)

      else if row.length > 1 and row.length is colLength
        rowFixed = filterFields(custumFn(addExtras(row)), colLength)
        query = queryFn(colFixed, rowFixed)

        if query
          console.log query

      else
        console.error "suspicious row count at row: #{index}", row

    process.exit(0)

  return



