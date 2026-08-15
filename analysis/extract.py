import openpyxl, pickle, datetime as dt
F='/root/.claude/uploads/c8c951a1-5f75-54ea-97b2-454c76055775/49421f7c-______V1.xlsm'
wb = openpyxl.load_workbook(F, data_only=True, read_only=True)
sheets={}
for name in ['始値','高値','安値','終値','出来高']:
    ws=wb[name]
    rows=list(ws.iter_rows(min_row=1, max_row=510, max_col=254, values_only=True))
    sheets[name]=rows
    print(name, len(rows), len(rows[0]))
pickle.dump(sheets, open('raw.pkl','wb'))
