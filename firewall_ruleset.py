#Code to mine ruleset
import pandas as pd

data=pd.read_excel("firewall_test.xlsx") #read data from Excel

print "\nFields are:{0}".format(list(data)) #to list headers

data.dropna(inplace=True) #to drop NaN
print data

#Count number of rules with field "Any"
no_fields_any=data[data.Source.str.contains("Any") | data.Destination.str.contains("Any") | data.Service.str.contains("Any")]
print "\nNumber of fields which contain Any:{0}".format(len(no_fields_any))
print "\nRows which contain Any\n"
print no_fields_any

#to write to excel
writer=pd.ExcelWriter('firewall_results.xlsx',engine='xlsxwriter')
no_fields_any.to_excel(writer, sheet_name='Sheet1')
writer.save()
