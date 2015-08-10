# converter
for reading excel files and then doing mysql inserts

# test:

coffee convert.coffee --file test.csv --table table 1>temp.sql

coffee convert.coffee --file test.ods --table table 1>temp.sql

# known issues:

the xls reader that I'm using has a documented bug where if data is the same in different fields then it doens't get processed, lol...
test.ods demonstrates this.