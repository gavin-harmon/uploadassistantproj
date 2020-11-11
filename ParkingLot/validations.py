
import os
import sys
import numpy as np
import pandas as pd
from cerberus import Validator

global mandatoryvalblanks, valmessage, valdf, vdata


def fetch_sdata():
    """Loads the input file. Note that the path has different logic for script vs exe execution.

    Before doing so it reads all the file names in "Submission" folder. currently the load process can only
    work with one file, but this may change as the project develops.
  """

    global sdata, spath

    if hasattr(sys, 'frozen'):
        spath = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Submission"))
    else:
        spath = os.path.join(os.path.dirname(__file__), "../Submission")

    print(spath)
    files = os.listdir(spath)

    """Get a list of only excel files in the path, find several extentions formats, case sensitive"""
    files = [files.lower() for files in files]
    files_xls = [f for f in files if f[-3:] in ('lsx', 'lsm', 'xls')]

    """     empty list to append to"""
    pathfiles = []

    """     create list of files with path"""
    for f in files_xls:
        makepathsfiles = os.path.join(str(spath), str(f))
        pathfiles.append(makepathsfiles)
    #    """     empty dataframe to append to"""
    #    df = pd.DataFrame()
    """     Read Summarize and append to df"""
    for f in pathfiles:
        global sdata
        sdata = pd.read_excel(f, sheet_name='Ptf_Monitoring_GROSS_Reins', na_values=[0], header=3,
                              converters={'Business Partner Name': str,
                                          'Type of Business': str, 'Type of Account': str, 'Distribution Type': str,
                                          'LOB': str, 'Distribution Channel': str,
                                          'Sub LOB': str, 'Business Partner ID Number': str, 'Product Name': str,
                                          'Product ID Number': str, 'Product Family': str,
                                          'Standard Product': str, })
        sdata.columns = sdata.columns.str.strip()

        """Remove rows with null business units"""
        sdata = sdata[sdata['Business Unit'].notnull()]
        return sdata


def fetch_manfields():
    """find mandatory fields"""
    """Import Template"""
    global manfields, tpath

    if hasattr(sys, 'frozen'):
        tpath = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Template"))
    else:
        tpath = os.path.join(os.path.dirname(__file__), "../Template")

    files = os.listdir(tpath)

    files = [files.lower() for files in files]

    """Get a list of only excel files in the path, find several extentions formats, case sensitive"""
    template = [f for f in files if f[-3:] in ('lsx', 'lsm', 'xls')]
    template = template[0]

    # template file path
    pathsfile = os.path.join(str(tpath), str(template))

    # empty dataframe
    mdf = pd.DataFrame()

    # read the lines with mandatory descriptions and fields
    mand = pd.read_excel(pathsfile, sheet_name='Ptf_Monitoring_GROSS_Reins', na_values=[0], header=0, nrows=3)
    mdf = mdf.append(mand)

    # Transpose and drop non field name data
    mdf = mdf.T.reset_index()

    # turn to a simple list rather than a pandas dataframe
    manfields = mdf[2][mdf[0] == "Mandatory"].values.tolist()



fetch_sdata()
fetch_manfields()
vdata=[]
vdata.append(sdata.replace(np.nan, '', regex=True))


mandatoryvalblanks = {}
valmessage = {}
valdf = {}
coldf = {}
for b in  manfields:
    # schema checks if blank
    schemamanfieldcheck = { b:  {'nullable': False, 'type': ['string', 'float', 'date'], 'empty': False  }, }

    # Cerberus functions
    v = Validator(schemamanfieldcheck)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation, make a dataframe with the field name in it
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    mandatoryvalblanks["{0}".format(b)] = pd.DataFrame(vdata.loc[rownums])

    if(len(mandatoryvalblanks["{0}".format(b)].index) == 0 ):
        v1= "Clear"
    else:
        valmessage["{0}".format(b)] = f'The following {len(mandatoryvalblanks["{0}".format(b)].index)} rows are missing entries for "{b}", please correct this and resubmit, or press "pass" and note why these rows are being submitted without product information'
        valdf["{0}".format(b)] = mandatoryvalblanks["{0}".format(b)]
        coldf["{0}".format(b)] = f'{b}'

#BR-xxx  Distribution Channel Blanks replaced by BR012

    # schema checks if blank
schema0 = { 'Distribution Channel': {'nullable': False,'type': 'string',  'empty': False},}

# Cerberus functions
v = Validator(schema0)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)
s0 = pd.DataFrame(vdata.loc[rownums])

if(len(s0.index) == 0 ):
    v0= "Clear"
else:
    print(f'The following {len(s0.index)} rows are missing Distribution Channel, please correct this and resubmit, or press "pass" and note why these rows are being submitted without product information')
    print(s0)

#BR-xxx Sub LOB Blanks replaced by BR0012

# schema checks if blank
schema1 = { 'Sub LOB': {'nullable': False, 'type': ['string'], 'empty': False  }, }

# Cerberus functions
v = Validator(schema1)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)
s1 = pd.DataFrame(vdata.loc[rownums])

if(len(s1.index) == 0 ):
    v1= "Clear"
else:
    print(f'The following {len(s1.index)} rows are missing Sub LOB, please correct this and resubmit, or press "pass" and note why these rows are being submitted without product information')
    print(s1)

#BR-xxx	B-Partner Blanks Blanks replaced by BR0012
# schema checks if blank
schema2 = { 'Business Partner Name': {'nullable': False,'type': 'string', 'empty': False}, }

# Cerberus functions
v = Validator(schema2)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)
s2 = pd.DataFrame(vdata.loc[rownums])

if(len(s2.index) == 0 ):
    v2= "Clear"
else:
    print(f'The following {len(s2.index)} rows have missing Business Partner Names, please correct this and resubmit, or press "pass" and note why these rows are being submitted without product information')
    print(s2)

#BR-xxx	Product Provided Blanks replaced by BR0012, maybe edit into a concatenation of all product fields
# schema checks if blank
schema3 = { 'Product Name': {'nullable': False,'type': 'string', 'empty': False}, }

# Cerberus functions
v = Validator(schema3)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)
s3 = pd.DataFrame(vdata.loc[rownums])

if(len(s3.index) == 0 ):
    v3= "Clear"
else:
    print(f'The following {len(s3.index)} rows have missing "Product Names", please correct this and resubmit, or press "pass" and note why these rows are being submitted without product information')
    print(s3)


vdata['Units of Risk'] = sdata['Units of Risk (Earned)'].add(sdata['Units of Risk (Written)'], fill_value=0).replace(np.nan, '', regex=True)


#BR016	Units of Risk Check (Earned or Written)
schema4 = { 'Units of Risk': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False, } }

# Cerberus functions
v = Validator(schema4)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)
s4 = pd.DataFrame(vdata.loc[rownums])

if(len(s4.index) == 0 ):
    v4= "Clear"
else:
    print(f'"Units of Risk" is not provided in the following {len(s4.index)} rows, please include them and resubmit or explain the local calculation for severity and anything preventing you from reporting it currently.')
    print(s4)



vdata['Number of Policies'] = sdata['Number of Policies (Earned)'].add(sdata['Number of Policies (Written)'], fill_value=0).replace(np.nan, '', regex=True)


#BR016	Number of Policies
# schema checks if greater than 0
schema5 = { 'Number of Policies': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False, } }

# Cerberus functions
v = Validator(schema5)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)

s5 = pd.DataFrame(vdata.loc[rownums])

if(len(s5.index) == 0 ):
    v5= "Clear"
else:
    print(f'"Number of Policies" is not provided in the following {len(s5.index)} rows, please include them and resubmit or explain the local calculation for frequency and anything preventing you from reporting it currently.')
    print(s5)


vdata['comsub'] = vdata['Commission Ratio'].multiply(1, fill_value=0).replace(np.nan, '', regex=True)

#BR020	Commission Ratio Check
# schema checks if greater than 100%
schema6 = { 'comsub': {'nullable': True, 'type': 'number', 'max': 1, 'required': False, } }

# Cerberus functions
v = Validator(schema6)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)

s6 = pd.DataFrame(vdata.loc[rownums])

if(len(s6.index) == 0 ):
    v6= "Clear"
else:
    print(f'"Commission Ratio is greater the "Earned Revenues Net of Taxes" in {len(s6.index)} rows, please reviuew and resubmit or explain these results with a note.')
    print(s6)



vdata['expsub'] = vdata['Expense Ratio'].multiply(1, fill_value=0).replace(np.nan, '', regex=True)

#BR019	Expense Ratio Check
# schema checks if greater than 100%
schema7 = { 'expsub': {'nullable': True, 'type': 'number', 'max': 1, 'required': False, } }

# Cerberus functions
v = Validator(schema7)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)

s7 = pd.DataFrame(vdata.loc[rownums])

if(len(s7.index) == 0 ):
    v7 = "Clear"
else:
    print(f'"Expense Ratio is greater the "Earned Revenues Net of Taxes" in {len(s7.index)} rows, please reviuew and resubmit or explain these results with a note.')
    print(s7)




vdata['Reporting Date From INT'] = pd.to_datetime(vdata['Reporting Date From']).dt.strftime("%Y%m%d").astype(int)

#BR006	Date Ranges
# schema checks if date is correct
schema8 = { 'Reporting Date From INT': {'type': ['number'] , 'max': 20190401 , 'min': 20190401 , 'required': False, } }

# Cerberus functions
v = Validator(schema8)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)

s8 = pd.DataFrame(vdata.loc[rownums])

if(len(s8.index) == 0 ):
    v8 = "Clear"
else:
    print(f'"Reporting Date From" is incorrect on {len(s8.index)} rows, the Reporting Period should start on June 1st, 2020.')
    print(s8)

vdata['Reporting Date To INT'] = pd.to_datetime(vdata['Reporting Date To']).dt.strftime("%Y%m%d").astype(int)

#BR006	Date Ranges
# schema checks if date is correct
schema9 = { 'Reporting Date To INT': {'type': ['number'] , 'max': 20200331 , 'min': 20200331 , 'required': False, } }

# Cerberus functions
v = Validator(schema9)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)

s9 = pd.DataFrame(vdata.loc[rownums])

if(len(s9.index) == 0 ):
    v9 = "Clear"
else:
    print(f'"Reporting Date To" is incorrect on {len(s9.index)} rows, the Reporting Period should end on May 31st, 2020.')
    print(s9)

vdata['Date of Analysis INT'] = pd.to_datetime(vdata['Date of Analysis']).dt.strftime("%Y%m%d").astype(int)

#BR006	Date Ranges
# schema checks if date is in correct range
schema10 = { 'Date of Analysis INT': {'type': ['number'] , 'max': 20200801 , 'min': 20200331 , 'required': False, } }

# Cerberus functions
v = Validator(schema10)
v.allow_unknown = True
v.require_all = True

# create a dictionary from the dataset (dict is a set where each item has an index)
# this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
data_dict = vdata.to_dict(orient='records')

# empty list to fill
rownums = []

# append a row to rownums with each record that fails validation
for idx, record in enumerate(data_dict):
    if not v.validate(record):
        rownums.append(idx)

s10 = pd.DataFrame(vdata.loc[rownums])

if(len(s10.index) == 0 ):
    v10 = "Clear"
else:
    print(f'"Date of Analysis" is incorrect on {len(s10.index)} rows, this should be after the the reporting period and before tomorrow.')
    print(s10)

vdata
print(manfields)

if __name__ == '__main__':
    valid
