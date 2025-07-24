uvicorn app.main:app --reload
python -m uvicorn app.main:app --reload
python -m fastapi dev main.py


Extract Tables

curl --location --request GET 'http://localhost:8000/extract?excel_path=C%3A%2FUsers%2Fanand.kumar%2FDocuments%2Fruleset%2FKEY-GR_PM.xlsx&start_sheet=End_Connection&end_sheet=Optional_Features' \
--header 'Content-Type: application/x-www-form-urlencoded' \
--data-urlencode 'excel_path=C:/Users/anand.kumar/Documents/ruleset/KEY-GR_PM.xlsx' \
--data-urlencode 'start_sheet=End_Connection' \
--data-urlencode 'end_sheet=Optional_Features'

Generate Ruletext

curl --location 'http://localhost:8000/generate-rules?excel_files_path=C%3A%2FUsers%2Fanand.kumar%2FDocuments%2Fruleset%2Ffiles'
