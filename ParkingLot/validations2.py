"""
Validations.py
====================================
This is the validation file that contains all validations.

*this is likely to be broken up in order to make th UX smoother.
*There is some
prework above the validations, and then each validation is separate in a section with the BRcode above it.
"""

"""Packages used"""

import os
import sys

import numpy as np
import pandas as pd
from cerberus import Validator

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
    vdata = sdata
    vdata['Units of Risk'] = sdata['Units of Risk (Earned)'].add(sdata['Units of Risk (Written)'],
                                                                 fill_value=0).replace(np.nan, '', regex=True)

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


"""cerberus is validation software. read more on this topic here   https://docs.python-cerberus.org/    """

""" validations section """

""" validations section """



global mandatoryvalblanks, valmessage, valdf, coldf, cleared, rowcounts

cleared=[]
mandatoryvalblanks={}
valmessage={}
valdf={}
coldf={}
rowcounts=[]

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
    rownums=[]

    # append a row to rownums with each record that fails validation, make a dataframe with the field name in it
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    mandatoryvalblanks["{0}".format(b)] = pd.DataFrame(vdata.loc[rownums])

    if len(mandatoryvalblanks["{0}".format(b)].index) == 0 :
        v0=f"{b}Clear"
        cleared.append(v0)
    else:
        if len(mandatoryvalblanks["{0}".format(b)].index) == 1:
            valmessage["{0}".format(b)] = f'The following row is missing an entry for "{b}", please correct this and resubmit,' \
                                          f' or press "pass" and note why this row is being submitted without product information'
        else:
            valmessage["{0}".format(b)]=f'The following {len(mandatoryvalblanks["{0}".format(b)].index)}' \
                                        f' rows are missing entries for "{b}", please correct this and resubmit,' \
                                        f' or press "pass" and note why these rows are being submitted without product information'
        rowcounts.append(len(mandatoryvalblanks["{0}".format(b)].index))
        valdf["{0}".format(b)]=mandatoryvalblanks["{0}".format(b)]
        coldf["{0}".format(b)]=f'{b}'



#BR-xxx  Distribution Channel Blanks replaced by BR012

#     # schema checks if blank
# schema0 = { 'Distribution Channel': {'nullable': False,'type': 'string',  'empty': False},}
#
# # Cerberus functions
# v = Validator(schema0)
# v.allow_unknown = True
# v.require_all = True
#
# # create a dictionary from the dataset (dict is a set where each item has an index)
# # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
# data_dict = vdata.to_dict(orient='records')
#
# # empty list to fill
# rownums = []
#
# # append a row to rownums with each record that fails validation
# for idx, record in enumerate(data_dict):
#     if not v.validate(record):
#         rownums.append(idx)
# s0 = pd.DataFrame(vdata.loc[rownums])
#
# if(len(s0.index)==0):
#     v0="Distribution Channel"
#     cleared.append(v0)
# else:
#     if len(s0.index)==1:
#         valmessage["Distribution Channel"] = f'The following row is missing an entry for "Distribution Channel",' \
#                                              f' please correct this and resubmit, or press "pass" and note why this' \
#                                              f' row is being submitted without "Distribution Channel"'
#     else:
#         valmessage["Distribution Channel"] = f'The following {len(s0.index)}' \
#                                       f' rows are missing entries for "Distribution Channel", please correct this and resubmit,' \
#                                       f' or press "pass" and note why these rows are being submitted without "Distribution Channel"'
#         rowcounts.append(len(s0.index))
#         valdf["Distribution Channel"]=s0
#         coldf["Distribution Channel"]="Distribution Channel"

#BR-xxx Sub LOB Blanks replaced by BR0012
#
# # schema checks if blank
# schema1={ 'Sub LOB': {'nullable': False, 'type': ['string'], 'empty': False  }, }
#
# # Cerberus functions
# v=Validator(schema1)
# v.allow_unknown = True
# v.require_all = True
#
# # create a dictionary from the dataset (dict is a set where each item has an index)
# # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
# data_dict = vdata.to_dict(orient='records')
#
# # empty list to fill
# rownums = []
#
# # append a row to rownums with each record that fails validation
# for idx, record in enumerate(data_dict):
#     if not v.validate(record):
#         rownums.append(idx)
# s1 = pd.DataFrame(vdata.loc[rownums])
#
# if(len(s1.index)==0 ):
#     v1='Sub LOB clear'
#     cleared.append(v1)
# else:
#     if len(s1.index) == 1:
#         valmessage['Sub LOB'] = f'The following row is missing an entry for "Sub LOB",' \
#                                              f' please correct this and resubmit, or press "pass" and note why this' \
#                                              f' row is being submitted without "Sub LOB"'
#     else:
#         valmessage["Sub LOB"] = f'The following {len(s1.index)}' \
#                                       f' rows are missing entries for "Sub LOB", please correct this and resubmit,' \
#                                       f' or press "pass" and note why these rows are being submitted without "Sub LOB"'
#         rowcounts.append(len(s1.index))
#         valdf["Sub LOB"]=s1
#         coldf["Sub LOB"]="Sub LOB"

#BR-xxx	B-Partner Blanks Blanks replaced by BR0012
# schema checks if blank
# schema2={ 'Business Partner Name': {'nullable': False,'type': 'string', 'empty': False}, }
#
# # Cerberus functions
# v=Validator(schema2)
# v.allow_unknown = True
# v.require_all = True
#
# # create a dictionary from the dataset (dict is a set where each item has an index)
# # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
# data_dict = vdata.to_dict(orient='records')
#
# # empty list to fill
# rownums = []
#
# # append a row to rownums with each record that fails validation
# for idx, record in enumerate(data_dict):
#     if not v.validate(record):
#         rownums.append(idx)
# s2 = pd.DataFrame(vdata.loc[rownums])
#
# if(len(s2.index)==0 ):
#     v2='Business Partner Name'
#     cleared.append(v2)
# else:
#     if len(s2.index) == 1:
#         valmessage['Business Partner Name'] = f'The following row is missing an entry for "Business Partner Name",' \
#                                              f' please correct this and resubmit, or press "pass" and note why this' \
#                                              f' row is being submitted without "Business Partner Name"'
#     else:
#         valmessage["Business Partner Name"] = f'The following {len(s2.index)}' \
#                                       f' rows are missing entries for "Business Partner Name", please correct this and resubmit,' \
#                                       f' or press "pass" and note why these rows are being submitted without "Business Partner Name"'
#         rowcounts.append(len(s2.index))
#         valdf["Business Partner Name"]=s2
#         coldf["Business Partner Name"]="Business Partner Name"

#BR-xxx	Product Provided Blanks replaced by BR0012, maybe edit into a concatenation of all product fields
# # schema checks if blank
# schema3 = { 'Product Name': {'nullable': False,'type': 'string', 'empty': False}, }
#
# # Cerberus functions
# v = Validator(schema3)
# v.allow_unknown = True
# v.require_all = True
#
# # create a dictionary from the dataset (dict is a set where each item has an index)
# # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
# data_dict = vdata.to_dict(orient='records')
#
# # empty list to fill
# rownums = []
#
# # append a row to rownums with each record that fails validation
# for idx, record in enumerate(data_dict):
#     if not v.validate(record):
#         rownums.append(idx)
# s3 = pd.DataFrame(vdata.loc[rownums])
#
# if(len(s3.index)==0 ):
#     v3='Product Name'
#     cleared.append(v3)
# else:
#     if len(s3.index) == 1:
#         valmessage['Product Name'] = f'The following row is missing an entry for "Product Name",' \
#                                              f' please correct this and resubmit, or press "pass" and note why this' \
#                                              f' row is being submitted without product information'
#     else:
#         valmessage["Product Name"] = f'The following {len(s3.index)}' \
#                                       f' rows are missing entries for "Product Name", please correct this and resubmit,' \
#                                       f' or press "pass" and note why these rows are being submitted without product information'
#         rowcounts.append(len(s3.index))
#         valdf["Product Name"]=s3
#         coldf["Product Name"]="Product Name"


# #BR016	Units of Risk Check (Earned or Written)
# vdata['Units of Risk'] = sdata['Units of Risk (Earned)'].add(sdata['Units of Risk (Written)'], fill_value=0).replace(np.nan, '', regex=True)
#
# schema4 = { 'Units of Risk': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False, } }
#
# # Cerberus functions
# v = Validator(schema4)
# v.allow_unknown = True
# v.require_all = True
#
# # create a dictionary from the dataset (dict is a set where each item has an index)
# # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
# data_dict = vdata.to_dict(orient='records')
#
# # empty list to fill
# rownums = []
#
# # append a row to rownums with each record that fails validation
# for idx, record in enumerate(data_dict):
#     if not v.validate(record):
#         rownums.append(idx)
# s4 = pd.DataFrame(vdata.loc[rownums])
#
# if(len(s4.index)==0 ):
#     v4='Units of Risk'
#     cleared.append(v4)
# else:
#     if len(s4.index) == 1:
#         valmessage['Units of Risk'] = f'The following row is missing an entry for "Units of Risk",' \
#                                              f' please correct this and resubmit, or press "pass" and note why this' \
#                                              f' row is being submitted without "Units of Risk"'
#     else:
#         valmessage["Units of Risk"] = f'The following {len(s4.index)}' \
#                                       f' rows are missing entries for "Units of Risk", please correct this and resubmit,' \
#                                       f' or press "pass" and note why these rows are being submitted without "Units of Risk"'
#         rowcounts.append(len(s4.index))
#         valdf["Units of Risk"]=s4
#         coldf["Units of Risk"]="Units of Risk"

# vdata['Number of Policies'] = sdata['Number of Policies (Earned)'].add(sdata['Number of Policies (Written)'], fill_value=0).replace(np.nan, '', regex=True)
#
# #BR016	Number of Policies
# # schema checks if greater than 0
# schema5 = { 'Number of Policies': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False, } }
#
# # Cerberus functions
# v = Validator(schema5)
# v.allow_unknown = True
# v.require_all = True
#
# # create a dictionary from the dataset (dict is a set where each item has an index)
# # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
# data_dict = vdata.to_dict(orient='records')
#
# # empty list to fill
# rownums = []
#
# # append a row to rownums with each record that fails validation
# for idx, record in enumerate(data_dict):
#     if not v.validate(record):
#         rownums.append(idx)
#
# s5 = pd.DataFrame(vdata.loc[rownums])
#
# if(len(s5.index)==0 ):
#     v5='Number of Policies'
#     cleared.append(v5)
# else:
#     if len(s5.index) == 1:
#         valmessage['Number of Policies'] = f'The following row is missing an entry for "Number of Policies",' \
#                                              f' please correct this and resubmit, or press "pass" and note why this' \
#                                              f' row is being submitted without "Number of Policies"'
#     else:
#         valmessage["Number of Policies"] = f'The following {len(s5.index)}' \
#                                       f' rows are missing entries for "Number of Policies", please correct this and resubmit,' \
#                                       f' or press "pass" and note why these rows are being submitted without "Number of Policies"'
#         rowcounts.append(len(s5.index))
#         valdf["Number of Policies"]=s5
#         coldf["Number of Policies"]="Number of Policies"
#

#BR020	Commission Ratio Check
vdata['comsub'] = vdata['Commission Ratio'].multiply(1, fill_value=0).replace(np.nan, '', regex=True)

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

if len(s6.index) == 0:
    v6='comsub'
    cleared.append(v6)
else:
    if len(s6.index) == 1:
        valmessage['comsub'] = f'The following row has a "Commission Ratio" over 100% of "Earned Revenue",' \
                                             f' please correct this and resubmit, or press "pass" and note why this' \
                                             f' row is being submitted with a high "Commission Ratio"'
    else:
        valmessage["comsub"] =f'The following {len(s6.index)} rows have a "Commission Ratio" over 100% of ' \
                              f'"Earned Revenue", please correct this and resubmit, or press "pass" and note ' \
                              f'why these rows are being submitted with a high "Commission Ratio"'

    rowcounts.append(len(s6.index))
    valdf["comsub"]=s6
    coldf["comsub"]="Commission Ratio"

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

if len(s7.index) == 0:
    v7 = 'expsub'
    cleared.append(v7)
else:
    if len(s7.index) == 1:
        valmessage['expsub'] = f'The following row has a "Expense Ratio" over 100% of "Earned Revenue",' \
                               f' please correct this and resubmit, or press "pass" and note why this' \
                               f' row is being submitted with a high "Expense Ratio"'
    else:
        valmessage["expsub"] = f'The following {len(s7.index)} rows have a "Expense Ratio" over 100% of ' \
                               f'"Earned Revenue", please correct this and resubmit, or press "pass" and note ' \
                               f'why these rows are being submitted with a high "Expense Ratio"'

    rowcounts.append(len(s7.index))
    valdf["expsub"] = s7
    coldf["expsub"] = "Expense Ratio"

schema8 = { 'Reporting Date From INT': {'type': ['number'] , 'max': 20200331 , 'min': 20200331 , 'required': False, } }

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

if (len(s8.index) == 0):
    v8 = 'Reporting Date From INT'
    cleared.append(v8)
else:
    if len(s8.index) == 1:
        valmessage['Reporting Date From INT'] = f'"Reporting Date From" is incorrect on this row, the Reporting Period ' \
                                                f'should start on June 1st, 2020. This is a critical validation, email ' \
                                                f'Global Portfolio Monitoring for instructions if you cannot provide the requested date range.'
    else:
        valmessage["Reporting Date From INT"] = f'"Reporting Date From" is incorrect on {len(s8.index)} rows, the Reporting Period ' \
                                                f'should start on June 1st, 2020. This is a critical validation, email ' \
                                                f'Global Portfolio Monitoring for instructions if you cannot provide the requested date range.'

    rowcounts.append(len(s8.index))
    valdf["Reporting Date From INT"] = s8
    coldf["Reporting Date From INT"] = "Reporting Date From INT"

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

if (len(s9.index) == 0):
    v9 = 'Reporting Date To INT'
    cleared.append(v9)
else:
    if len(s9.index) == 1:
        valmessage['Reporting Date To INT'] = f'"Reporting Date To" is incorrect on this row, the Reporting Period should' \
                                              f' end on May 31st, 2020. This is a critical validation, email Global Portfolio' \
                                              f' Monitoring for instructions if you cannot provide the requested date range.'
    else:
        valmessage["Reporting Date To INT"] = f'"Reporting Date To" is incorrect on {len(s9.index)} rows, the Reporting Period' \
                                              f' should end on May 31st, 2020. This is a critical validation, email Global Portfolio' \
                                              f' Monitoring for instructions if you cannot provide the requested date range.'

    rowcounts.append(len(s9.index))
    valdf["Reporting Date To INT"] = s9
    coldf["Reporting Date To INT"] = "Reporting Date To INT"


#BR006	Date Ranges
#schema checks if date is in correct range
vdata['Date of Analysis INT'] = pd.to_datetime(vdata['Date of Analysis']).dt.strftime("%Y%m%d").astype(int)

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

if len(s10.index) == 0:
    v10 = 'Date of Analysis INT'
    cleared.append(v10)
else:
    if len(s10.index) == 1:
        valmessage['Date of Analysis INT'] = f'"Date of Analysis" is incorrect on this row, this should be after ' \
                                             f'the reporting period and before tomorrow.'
    else:
        valmessage['Date of Analysis INT'] = f'"Date of Analysis" is incorrect on {len(s10.index)} rows, this should be after ' \
                                             f'the reporting period and before tomorrow.'
    rowcounts.append(len(s10.index))
    valdf['Date of Analysis INT'] = s10
    coldf['Date of Analysis INT'] = 'Date of Analysis INT'

#BR013	Unaggregated Attributes

"""
create a field to evaluate
"""
vdata['Selected Fields for Duplicates'] = vdata['Type of Business'].str.strip().fillna("")+vdata['Type of Account'].str.strip().fillna("")\
                                          +vdata['Distribution Type'].str.strip().fillna("")+vdata['LOB'].str.strip().fillna("")\
                                          +vdata['Distribution Channel'].str.strip().fillna("")+vdata['Sub LOB'].str.strip().fillna("")\
                                          +vdata['Business Partner Name'].str.strip().fillna("")+vdata['Business Partner ID Number'].str.strip().fillna("")\
                                          +vdata['Product Name'].str.strip().fillna("")+vdata['Product ID Number'].str.strip().fillna("")\
                                          +vdata['Product Family'].str.strip().fillna("")+vdata['Standard Product'].str.strip().fillna("")

vcounts = pd.DataFrame(vdata['Selected Fields for Duplicates'].value_counts()).\
    reset_index().rename(columns={'Selected Fields for Duplicates' : 'Duplicate Count'})

vdata = vdata.merge(vcounts, how= 'left',  left_on='Selected Fields for Duplicates', right_on='index', suffixes=(False, False))

schema11 = {'Duplicate Count': {'type': ['number'], 'max': 1,  'required': False, }}

# Cerberus functions
v = Validator(schema11)
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

s11 = pd.DataFrame(vdata.loc[rownums])

if len(s11.index) == 0:
    v11 = 'Duplicate Count'
    cleared.append(v11)
else:
    valmessage['Duplicate Count'] = f'These combinations of Distribution Channel, Sub LOB, B-Partner, and Product are repeated,'\
                                  f' please aggregate and resubmit, or pass and notate.'
    rowcounts.append(len(s11.index))
    valdf['Duplicate Count'] = s11
    coldf['Duplicate Count'] = 'Duplicate Count'

vdata
""" end validations section """
