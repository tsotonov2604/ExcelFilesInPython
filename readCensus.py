import openpyxl, pprint
print('Opening workbook...')
wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb.active
countryData = {}

print('Reading rows...')
for row in range(2,sheet.max_row + 1):
    state = sheet['B'+str(row)].value
    country = sheet['C'+ str(row)].value
    pop = sheet['D'+ str(row)].value

    #make sure that the key for this state exists 
    countryData.setdefault(state,{})
    #make sure tht the ket for this country in this state exists
    countryData[state].setdefault(country,{'tracts':0,'pop':0})
    #each row represents one sensus tract so incrament by one
    countryData[state][country]['tracts'] += 1
    #increase the country pop by the pop in this census tract
    countryData[state][country]['pop'] += int(pop)
    
print('Writing results...')
results = open('census2010.py','w')
results.write('allData = ' + pprint.pformat(countryData))
results.close()
print('Done')
