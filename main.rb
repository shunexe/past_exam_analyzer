require "google_drive"
require "googleauth"
 
session = GoogleDrive::Session.from_config("config.json")
# https://docs.google.com/spreadsheets/d//
# 事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
spreadsheet = session.spreadsheet_by_key("1xdS7RJqlP-VpPgdYu63GLgbYJAXw0CDYMudP_qNZNJo")
allSheets = spreadsheet.worksheets

items = []
pref =[]
allSheets.each do |sheet|
  unless(sheet.title.include?("H"))
    next
  end
  result = {}
  count = 1
  for num in 1 .. sheet.rows.length #これは年ごとに異なる
    pp sheet[num,3].split("，") 
    items.push(sheet[num,3].split("，")).flatten!
    unless sheet[num,1]==""
      pref << sheet[num,1]
    end
  end
end
pref_result = pref.group_by(&:itself).map{ |key,value| [key,value.count]}
allResults = items.group_by(&:itself).map{ |key,value| [key,value.count]}
result_sheet = spreadsheet.worksheet_by_title("result")
result_sheet.update_cells(1,4,allResults.sort { |a, b| b[1] <=> a[1] })
result_sheet.update_cells(1,1,pref_result.sort { |a, b| b[1] <=> a[1] })
result_sheet.save