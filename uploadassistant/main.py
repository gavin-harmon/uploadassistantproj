"""
This this module creates the Upload Assistant. See the READ ME for instructions on how to deploy it into a Windows exe
application

Any changes to this document need to tested to 1) execute in the console 2) execute from the terminal
and 3) have a functional, shareable .exe output via pyinstaller

===================

"""
# libraries and modules used

import os
import sys # these two allow python to make system commands like run
import win32com.client as win32 # to open outlook
from datetime import datetime # for times
import mouse #interacts with a users mouse
import time # more time
import numpy as np # stats package
import pandas as pd # data package
import glob #finds all the pathnames matching a specified pattern
import json #reads and exports json
import tkinter as tk #GUI package
from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image #pillow image utilities for tkinter
from pandastable import Table, TableModel #pillow image utilities for tkinter
import importlib# part of the outlook email process
from cerberus import Validator# runs validation section
import nicexcel as nl# output excel in report format
import zipfile
from zipfile import ZipFile # make zip files
from os.path import basename #file path references

"""data import modules"""
global vdata, sdata, comlist, vmessagelist, sublist, vn, collist, cleared, rownums, vdflist, vlock, spath, bgimage, instimage, reports_dict, reports_dict_float
# set variable

"""controls the working directory for when being processed by documentation programs and executable builders"""
if hasattr(sys, 'frozen'):
    os.chdir(os.path.dirname(os.path.realpath(sys.executable)))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))


path = os.getcwd().replace("dist", "")
#main directory

spath = os.path.join( path, "Submission")
#Submission folder


tpath = os.path.join( path, "Template")
#Template folder

bgimage = os.path.join( path, "Images/bg5.png")
#Import Background

attachpath = os.path.join(path, "Output")
attachzip = os.path.join(path, "Output.zip")
distzip = os.path.join(path, "dist/Output.zip")
# create a ZipFile object




llock = {1}
vn=0
vlock = {}
# empty set to deactivate button response
# if the script has been run on the screen, this will be locked and not run again, the user pressing "back" unlocks
# this by setting llock={}

def lock():
    """This is called at various points, usually after actions, so that they do not repeat. If len(llock) ==1 then
    do nothing."""
    global llock
    llock = {1}


def unlock(self):
    """This is called at various points, usually after actions, to allow processes to occur."""
    global llock
    llock = {}

"""Import Submission"""

def fetch_sdata():
    """Loads the input file. Note that the path has different logic for script vs exe execution.

    Before doing so it reads all the file names in "Submission" folder. currently the load process can only
    work with one file, but this may change as the project develops.
  """
    if len(llock) == 0:
        global sdata, spath

        files_xls = glob.glob(os.path.join(spath, f'*.xls*'))

        try:
            recent_vers = max(files_xls, key=os.path.getctime)
        except ValueError:
            tk.messagebox.showerror(title="Error", message=f'Please check the last saved submission file in {spath}. '
                   f'This should be an Excel file and the data on a sheet named "Ptf_Monitoring_GROSS_Reins".')
            os.startfile(spath)
            sys.exit()

        """     empty list to append to"""

        subfile = os.path.join(str(spath), str(recent_vers))

        """     Read Summarize and append to df"""

        global sdata

        try:
            sdata = pd.read_excel(subfile, sheet_name='Ptf_Monitoring_GROSS_Reins', na_values=[''], header=3,
                                  converters={'Business Partner Name': str,
                                              'Type of Business': str, 'Type of Account': str, 'Distribution Type': str,
                                              'LOB': str, 'Distribution Channel': str,
                                              'Sub LOB': str, 'Business Partner ID Number': str, 'Product Name': str,
                                              'Product ID Number': str, 'Product Family': str,
                                              'Standard Product': str, })
        except PermissionError:
            tk.messagebox.showerror(title="Error", message="The submission file seems to be open in Excel. Please "
                                                           "save and close the file and restart the application.")
            sys.exit()

        sdata.columns = sdata.columns.str.strip()

        """Remove rows with null business units"""
        sdata = sdata[sdata['Business Unit'].notnull()]

        try:
            sdata['Country'].fillna(sdata['Business Unit'])
        except KeyError:
            sdata.insert(1,'Country',sdata['Business Unit'])

        return sdata

    else:
        lock()

def check_sheaders():
    """
    Check the headers of sdata. This is a validation that the headers are correct. it is not in the final 2020Q3
    request version.
    """
    global sdata, headers, sheaders, missingheaders, badheaders, headererrors, vdata

    sheaders = pd.Series(sdata.columns.values.tolist(), dtype="object")
    headers = pd.Series(headers)
    headers = headers.append(pd.Series('Identifier to pull in results from Data tab'), ignore_index=True, verify_integrity=False)
    try:
        if headers == sheaders:
            a=1
        else:
            a=2

    except ValueError:
        missingheaders = headers[~headers.isin(sheaders)]
        badheaders = sheaders[~sheaders.isin(headers)]

        mydict = {'These header(s) are missing': np.array(missingheaders), 'These header(s) do not conform to the template': np.array(badheaders)}

        headererrors = pd.DataFrame({ key:pd.Series(value) for key, value in mydict.items() })

def fetch_headers():
    """Loads headers from Template."""

    if len(llock) == 0:
        global headers, manfields, tpath

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
        headers = mdf[2][mdf[0] != ""].values.tolist()
        manfields = mdf[2][mdf[0] == "Mandatory"].values.tolist()

    else:
        lock()


"""     empty list to append to, eash time fetch each uploaded DataFrame will be a member of
 this list vdata[0], vdata[1] and so on. Reference the lates upload always as vdata[-1]"""

global sdata, manfields

vdata = []

""" Report Generation Section"""

def make_reports(df):
    """This generates the calculated fields in aggregate for each report view."""
    # This is a list of report views with thier attributes ex "Report name":("Attribute 1", "Attribute 2")
    global reports_dict, reports_dict_float, sdata

    rep_attrs_dict = ({"Country": ["Country", "Currency"], "Distribution Channel": ["Country", "Currency", "Distribution Channel"],
                       "Sub Line of Business": ["Country", "Currency", "Sub LOB", ],
                       "Business Partner": ["Country", "Currency", "Business Partner Name", "Business Partner ID Number", ],
                       "Product": ["Country", "Currency", "Sub LOB", "Product Name", "Product ID Number", ],
                       "B Partner & Product Combo": ["Country",  "Currency", "Business Partner Name",
                        "Business Partner ID Number", "Sub LOB", "Product Name", "Product ID Number",]})

    report_base_cols = ["Written Revenues net of Taxes", "Earned Revenues net of Taxes",
                        "Actual Incurred Losses (Paid + OCR + IBNR)",
                        "Earned Base Commissions", "Earned Over-Commissions", "Upfront Cash Payments", "Total Expenses",
                        "Number of Persons Involved in Claims (Paid + OCR + IBNR)", "Units of Risk (Earned)",
                        "Contribution Margin - HQ View", "Profit or Loss", ]
    rep_all_cols = []

    for key in rep_attrs_dict:
        item = rep_attrs_dict[key].copy()
        for i in report_base_cols:
            item.append(i)
        rep_all_cols.append(item)

    rep_rawdfs = []

    for i in rep_all_cols:
        s_df = sdata[i].copy().fillna(0)
        rep_rawdfs.append(s_df)

    rep_dfs = []

    for key, i in zip(rep_attrs_dict, rep_rawdfs):
        group_df = i.groupby(rep_attrs_dict[key], as_index=False).sum()
        sort_df = group_df.sort_values(by="Earned Revenues net of Taxes", ascending=False)
        rep_dfs.append(sort_df)

    for i in rep_dfs:
        i["Loss Ratio"] = i["Actual Incurred Losses (Paid + OCR + IBNR)"].values / i["Earned Revenues net of Taxes"].values
        i["Loss Ratio"] = i["Loss Ratio"].fillna(0)
        i["Loss Ratio Float"] = i["Loss Ratio"].values
        i["Loss Ratio"] = pd.Series(["{0:.1f}%".format(val * 100) for val in i["Loss Ratio"]],
                                    index=i["Loss Ratio"].index)

        i["Commission Ratio"] = ((i["Earned Base Commissions"].values + i["Earned Over-Commissions"].values
                                  + i["Upfront Cash Payments"].values) / i["Earned Revenues net of Taxes"].values)
        i["Commission Ratio"] = i["Commission Ratio"].fillna(0)
        i["Commission Ratio Float"] = i["Commission Ratio"].values
        i["Commission Ratio"] = pd.Series(["{0:.1f}%".format(val * 100) for val in i["Commission Ratio"]],
                                          index=i["Commission Ratio"].index)

        i["Expense Ratio"] = i["Total Expenses"].values / i["Earned Revenues net of Taxes"].values
        i["Expense Ratio"] = i["Expense Ratio"].fillna(0)
        i["Expense Ratio Float"] = i["Expense Ratio"].values
        i["Expense Ratio"] = pd.Series(["{0:.1f}%".format(val * 100) for val in i["Expense Ratio"]],
                                       index=i["Expense Ratio"].index)

        i["Contribution Margin"] = i["Contribution Margin - HQ View"].values / i["Earned Revenues net of Taxes"].values
        i["Contribution Margin"] = i["Contribution Margin"].fillna(0)
        i["Contribution Margin Float"] = i["Contribution Margin"].values
        i["Contribution Margin"] = pd.Series(["{0:.1f}%".format(val * 100) for val in i["Contribution Margin"]],
                                             index=i["Contribution Margin"].index)

        i["Combined Ratio"] = (
                (i["Actual Incurred Losses (Paid + OCR + IBNR)"].values + i["Earned Base Commissions"].values
                 + i["Earned Over-Commissions"].values + i["Upfront Cash Payments"].values
                 + i["Total Expenses"].values) / i["Earned Revenues net of Taxes"].values)
        i["Combined Ratio"] = i["Combined Ratio"].fillna(0)
        i["Combined Ratio Float"] = i["Combined Ratio"].values
        i["Combined Ratio"] = pd.Series(["{0:.1f}%".format(val * 100) for val in i["Combined Ratio"]],
                                        index=i["Combined Ratio"].index)

        # Frequency
        i["Frequency"] = i["Number of Persons Involved in Claims (Paid + OCR + IBNR)"].values / \
                         i["Units of Risk (Earned)"].values
        i["Frequency"] = i["Frequency"].fillna(0)
        i["Frequency Float"] = i["Frequency"].values
        i["Frequency"] = pd.Series(["{0:.1f}%".format(val * 100) for val in i["Frequency"].values], index=i["Frequency"].index)

        i["Severity"] = i["Actual Incurred Losses (Paid + OCR + IBNR)"].values / \
                        i["Number of Persons Involved in Claims (Paid + OCR + IBNR)"].values
        i["Severity"] = i["Severity"].fillna(0)
        i["Severity Float"] = i["Severity"].values
        i["Severity"] = pd.Series(["{:,.0f}".format(val) for val in i["Severity"]], index=i["Severity"].index)

        i["Written Revenues net of Taxes Float"] = i["Written Revenues net of Taxes"].values
        i["Written Revenues net of Taxes"] = pd.Series(
            ["{:,.0f}".format(val) for val in i["Written Revenues net of Taxes"]],
            index=i["Written Revenues net of Taxes"].index).fillna(0)
        i["Earned Revenues net of Taxes Float"] = i["Earned Revenues net of Taxes"].values
        i["Earned Revenues net of Taxes"] = pd.Series(
            ["{:,.0f}".format(val) for val in i["Earned Revenues net of Taxes"]],
            index=i["Earned Revenues net of Taxes"].index).fillna(0)
        i["Profit or Loss Float"] = i["Profit or Loss"].values
        i["Profit or Loss"] = pd.Series(["{:,.0f}".format(val) for val in i["Profit or Loss"]],
                                        index=i["Profit or Loss"].index)

    rep_val_cols = ["Written Revenues net of Taxes", "Earned Revenues net of Taxes", "Loss Ratio", "Commission Ratio",
                    "Expense Ratio", "Contribution Margin",
                    "Combined Ratio", "Profit or Loss", "Frequency", "Severity"]

    rep_val_cols_float = ["Written Revenues net of Taxes Float", "Earned Revenues net of Taxes Float", "Loss Ratio Float", "Commission Ratio Float",
                    "Expense Ratio Float", "Contribution Margin Float",
                    "Combined Ratio Float", "Profit or Loss Float", "Frequency Float", "Severity Float"]

    rep_cols = []

    for key in rep_attrs_dict:
        item = rep_attrs_dict[key].copy()
        for i in rep_val_cols:
            item.append(i)
        rep_cols.append(item)

    reports_dict = {}

    for key, i, j in zip(rep_attrs_dict, rep_dfs, rep_cols):
        reports_dict.update({key: i[j]})

    rep_cols_float = []

    for key in rep_attrs_dict:
        item = rep_attrs_dict[key].copy()
        for i in rep_val_cols_float:
            item.append(i)
        rep_cols_float.append(item)

    reports_dict_float = {}

    for key, i, j in zip(rep_attrs_dict, rep_dfs, rep_cols_float):
        reports_dict_float.update({key: i[j]})


""" End Report Generation Section"""

""" validations section """

def valid(vdata, manfields, sdata):
    """This section runs the validations logic"""
    global validationerrors, valmessage, valdf, coldf, subtitle, cleared, rowcounts

    cleared=[]
    validationerrors={}
    valmessage={}
    valdf={}
    coldf={}
    subtitle={}
    rowcounts=[]
    dict_list=[]
    dict_list_float=[]
    repnames=[]

    print("begin vals")

    # hardcoded to avoid user interface
    valid_field_values_dict = {'Business Unit': np.array(
        ['AT', 'AU', 'BE', 'BR', 'CA', 'CH', 'CN', 'CZ', 'DE', 'ES', 'FOS', 'FR', 'GR', 'IN', 'IT', 'JP', 'ME', 'MX',
           'NL', 'NZ', 'PL', 'PT', 'RU', 'SG', 'TH', 'TU', 'UK', 'US']), 'Country': np.array(['AT', 'AU', 'BE', 'BR',
           'CA', 'CH', 'CN', 'CZ', 'DE', 'DK', 'EE', 'ES', 'FI', 'FOS', 'FR', 'GR', 'IN', 'IS', 'IT', 'JP', 'LT', 'LV',
           'ME', 'MX', 'NL', 'NO', 'NZ', 'PL', 'PT', 'RU','SC', 'SE', 'SG', 'TH', 'TU', 'UK', 'US']), 'Currency': np.array(
           ['AED', 'ALL','ARS' ,'ATS' ,'AUD' ,'BAM' ,'BEF' ,'BGN' ,'BHD' ,'BND' ,'BRL', 'CAD' ,'CHF' ,'CLP' ,'CNY',
           'COP', 'CRC' ,'CYP' ,'CZK' ,'DEM' ,'DKK', 'DOP', 'EEK', 'EGP', 'ESP', 'EUR', 'FIM', 'FRF', 'GBP', 'GHS',
           'GRD', 'GTQ', 'HKD', 'HRK', 'HUF', 'IDR', 'IEP', 'ILS', 'INR', 'ISK', 'ITL', 'JOD', 'JPY', 'KES', 'KRW',
           'KWD', 'LBP', 'LKR', 'LTL', 'LUF', 'LVL', 'MAD', 'MDL', 'MGA', 'MOP', 'MTL', 'MUR', 'MXN', 'MYR', 'NAD',
           'NLG', 'NOK', 'NZD', 'OMR', 'PEN', 'PHP', 'PLN', 'PTE', 'QAR', 'RON', 'RSD', 'RUB', 'SAR', 'SEK', 'SGD',
           'SIT', 'SKK', 'THB', 'TND', 'TRY', 'TWD', 'TZS', 'UAH', 'UGX', 'USD', 'VND', 'XOF', 'XPF', 'ZAR']),
           'Region':np.array(['APAC', 'North America', 'North, Central & Eastern Europe', 'Western Europe, LATAM & MEA',
            ]), 'Type of Analysis':np.array([ "Most Recently 12 Months",  "Year To Date"]), 'Type of Business':
            np.array([ 'Insurance', 'Reinsurance', 'Service', ]),  'Type of Account': np.array(['Global - FOE',
            'Global - FOS', 'Local']), 'Distribution Type': np.array(['B2B', 'B2B2C', 'B2C']), 'Distribution Channel':
            np.array(['Allianz Affiliations', 'Banks / Credit Cards', 'Brokers', 'Carriers - Airlines',
            'Carriers - Cruise, Ferry', 'Carriers - Train, Bus', 'Direct (Allianz Partners)', 'Event', 'Lodging',
            'Managing General Agents', 'Offline Tour Operators', 'Offline Travel Agencies', 'Online Tour Operators',
            'Online Travel Agencies (OTAs)', 'Other Niche Travel Market', 'Other non-Allianz Insurers',
            'Payment Administrators', 'Schools / Universities', 'Visa Centers']), 'LOB': np.array(['Travel']),
            'Sub LOB': np.array([ 'Collision Damage Waiver (CDW)', 'Corporate Travel', 'Event Ticket Cancellation',
            'Expatriates', 'Standalone Cancellation - Multi-Trip', 'Standalone Cancellation - Single Trip',
            'TPA/Claims Handling / Service Only Products', 'Travel Medical Long term - Single Trip',
            'Travel Medical Multi Trip (Annual Insurance)', 'Travel Medical Short Term - Single Trip',
            'Travel Package w/o Cancellation Long Term Single Trip',
            'Travel Package w/o Cancellation Multi Trip (Annual Insurance)',
            'Travel Package w/o Cancellation Short Term Single Trip',
            'Travel Package with Cancellation Long Term Single Trip',
            'Travel Package with Cancellation Multi Trip (Annual Insurance)',
             'Travel Package with Cancellation Short Term Single Trip', 'Tuition','Other']), }

    valid_field_values = pd.DataFrame({key: pd.Series(value) for key, value in valid_field_values_dict.items()})

    valid_field_list = list(valid_field_values_dict.keys())

    melted = pd.melt(valid_field_values)

    bad_values = []
    all_bad_values = []

    for idx, i in enumerate(valid_field_list):
        all_bad_values.append(pd.DataFrame(vdata[i][~vdata[i].str.upper().isin(melted['value'].str.upper())]))
        bad_values.append(pd.melt(all_bad_values[idx]))

    bad_values = pd.concat(bad_values).drop_duplicates()

    bad_values = bad_values.reset_index(drop=True)

    messages = []

    for idx, i in enumerate(bad_values.T):
        message = f'"{bad_values["value"][idx]}" is not a valid value for field "{bad_values["variable"][idx]}"; ' \
                  f'correct in submission or explain in column "Notes".'
        messages.append(message)

    messages = pd.DataFrame(messages, columns=['Field Value Errors'])


    if(len(messages)==0):
        v0="Validated Fields"
        cleared.append(v0)
    else:
        valmessage["Validated Fields"] = f'Expand the column to read the entries that do not ' \
                                    f'match the appropriate dropdown selections. Please determine the precise name' \
                                    f' to replace the value with and resubmit.'
        rowcounts.append(len(messages))
        valdf["Validated Fields"]=messages
        coldf["Validated Fields"]="Validated Fields"
        subtitle["Validated Fields"] = "Value Match Check"

    summschema0 = {'comsub': {'nullable': True,  'min': -.00000001, 'max': 1, 'required': False, }}
    print("summschema0")
    for idx, key in enumerate(reports_dict):
        dict_list.append(reports_dict[key].to_dict(orient='records'))

    for idxx, key in enumerate(reports_dict_float):

        reports_dict_float[key]['comsub'] = reports_dict_float[key]['Commission Ratio Float'].multiply(1,
                                                                        fill_value=0).replace(np.nan, '', regex=True)
        dict_list_float.append(reports_dict_float[key].to_dict(orient='records'))
        repnames.append(key)

        # schema checks if greater than 100%

        # Cerberus functions
        v = Validator(summschema0)
        v.allow_unknown = True
        v.require_all = True

        # create a dictionary from the dataset (dict is a set where each item has an index)
        # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
        # empty list to fill
        rownums = []

        # append a row to rownums with each record that fails validation
        for idx, record in enumerate(dict_list_float[idxx]):
            if not v.validate(record):
                rownums.append(idx)
        sa0 = pd.DataFrame(dict_list[idxx]).loc[rownums]

        if len(sa0.index) == 0:
            cleared.append(f'comsub{repnames[idxx]}')
        else:
            if len(sa0.index) == 1:
                valmessage[
                    f'comsub{repnames[idxx]}'] = f'The following row should have a non-negative "Commission Ratio" ' \
                                                 f'less than 100% of "Earned Revenue"' \
                                                 f' at the {repnames[idxx]} level; please correct this and resubmit, '\
                                                 f'or press "pass" and note why this' \
                                                 f' row is being submitted with the presented "Commission Ratio".'
            else:
                valmessage[
                    f'comsub{repnames[idxx]}'] = f'The following {len(sa0.index)} rows should have a non-negative' \
                                            f' "Commission Ratio" less than 100% of ' \
                                            f'"Earned Revenue" at the {repnames[idxx]} level; please correct ' \
                                            f'this and resubmit, or press "pass" and note ' \
                                            f'why these rows are being submitted with the presented "Commission Ratio".'

            rowcounts.append(len(sa0.index))
            valdf[f'comsub{repnames[idxx]}'] = sa0
            coldf[f'comsub{repnames[idxx]}'] = "Commission Ratio"
            subtitle[f'comsub{repnames[idxx]}'] = "Aggregate Result Check"


    summschema1 = {'expsub': {'nullable': True, 'min': -.00000001, 'max': 1, 'required': False, }}
    print("summschema1")
    for idx, key in enumerate(reports_dict):
        dict_list.append(reports_dict[key].to_dict(orient='records'))

    for idxx, key in enumerate(reports_dict_float):

        reports_dict_float[key]['expsub'] = reports_dict_float[key]['Expense Ratio Float']
        dict_list_float.append(reports_dict_float[key].to_dict(orient='records'))
        repnames.append(key)

        # schema checks if greater than 100%

        # Cerberus functions
        v = Validator(summschema1)
        v.allow_unknown = True
        v.require_all = True

        # create a dictionary from the dataset (dict is a set where each item has an index)
        # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
        # empty list to fill
        rownums = []

        # append a row to rownums with each record that fails validation
        for idx, record in enumerate(dict_list_float[idxx]):
            if not v.validate(record):
                rownums.append(idx)
        sa1 = pd.DataFrame(dict_list[idxx]).loc[rownums]

        if len(sa1.index) == 0:
            cleared.append(f'expsub{repnames[idxx]}')
        else:
            if len(sa1.index) == 1:
                valmessage[
                    f'expsub{repnames[idxx]}'] = f'The following row should have a non-negative "Expense Ratio" less '\
                                                 f'than 100% of "Earned Revenue"' \
                                                 f' at the {repnames[idxx]} level; please correct this and ' \
                                                 f'resubmit, or press "pass" and note why this' \
                                                 f' row is being submitted with the presented "Expense Ratio".'
            else:
                valmessage[
                    f'expsub{repnames[idxx]}'] = f'The following {len(sa1.index)} rows should have a non-negative' \
                                            f' "Expense Ratio" less than 100% of ' \
                                            f'"Earned Revenue" at the {repnames[idxx]} level; please correct' \
                                            f' this and resubmit, or press "pass" and note ' \
                                            f'why these rows are being submitted with the presented "Expense Ratio".'

            rowcounts.append(len(sa1.index))
            valdf[f'expsub{repnames[idxx]}'] = sa1
            coldf[f'expsub{repnames[idxx]}'] = "Expense Ratio"
            subtitle[f'expsub{repnames[idxx]}'] = "Aggregate Result Check"

    summschema2 = {'losssub': {'nullable': True, 'min': 0,   'required': False, }}
    print("summschema2")

    for idx, key in enumerate(reports_dict):
        dict_list.append(reports_dict[key].to_dict(orient='records'))

    for idxx, key in enumerate(reports_dict_float):

        reports_dict_float[key]['losssub'] = reports_dict_float[key]['Loss Ratio Float']
        dict_list_float.append(reports_dict_float[key].to_dict(orient='records'))
        repnames.append(key)

        # schema checks if greater than 100%

        # Cerberus functions
        v = Validator(summschema2)
        v.allow_unknown = True
        v.require_all = True

        # create a dictionary from the dataset (dict is a set where each item has an index)
        # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
        # empty list to fill
        rownums = []

        # append a row to rownums with each record that fails validation
        for idx, record in enumerate(dict_list_float[idxx]):
            if not v.validate(record):
                rownums.append(idx)
        sa2 = pd.DataFrame(dict_list[idxx]).loc[rownums]

        if len(sa2.index) == 0:
            cleared.append(f'losssub{repnames[idxx]}')
        else:
            if len(sa2.index) == 1:
                valmessage[
                    f'losssub{repnames[idxx]}'] = f'The following row should have a non-negative "Loss Ratio" ' \
                                                 f' at the {repnames[idxx]} level; please correct this and resubmit,' \
                                                  f' or press "pass" and note why this' \
                                                 f' row is being submitted with a negative "Loss Ratio".'
            else:
                valmessage[
                    f'losssub{repnames[idxx]}'] = f'The following {len(sa2.index)} rows should have a non-negative' \
                                                  f' "Loss Ratio" ' \
                                                 f'at the {repnames[idxx]} level; please correct this and resubmit, ' \
                                                  f'or press "pass" and note ' \
                                                 f'why these rows are being submitted with a negative "Loss Ratio".'

            rowcounts.append(len(sa2.index))
            valdf[f'losssubsub{repnames[idxx]}'] = sa2
            coldf[f'losssubsub{repnames[idxx]}'] = "Loss Ratio"
            subtitle[f'losssubsub{repnames[idxx]}'] = "Aggregate Result Check"

    print("manfields start")
    for b in manfields:
        """
        This checks each mandatory field on a loop for blank values at the granular level.
        """
        # schema checks if blank
        schemamanfieldcheck = { b:  {'nullable': False, 'type': ['string', 'float', 'date'], 'empty': False  }, }

        # Cerberus functions
        v = Validator(schemamanfieldcheck)
        v.allow_unknown = True
        v.require_all = True
        vdata["concat2"] = vdata["Country"] + vdata["Currency"]
        vvdata = pd.DataFrame(vdata[["concat2", b]])
        # create a dictionary from the dataset (dict is a set where each item has an index)
        data_dict = vvdata.to_dict(orient='records')

        # empty list to fill
        rownums=[]

        # append a row to rownums with each record that fails validation, make a dataframe with the field name in it
        for idx, record in enumerate(data_dict):
            if not v.validate(record):
                rownums.append(idx)

        validationerrors["{0}".format(b)] = pd.DataFrame(vdata.loc[rownums])

        if len(validationerrors["{0}".format(b)].index) == 0:
            v0=f"{b}Clear"
            cleared.append(v0)
        else:
            if len(validationerrors["{0}".format(b)].index) == 1:
                valmessage["{0}".format(b)] = f'The following row is missing an entry for "{b}"; please correct' \
                                              f' this and resubmit,' \
                                              f' or press "pass" and note why this row is being submitted with' \
                                              f' blank entries for "{b}".'
            else:
                valmessage["{0}".format(b)] = f'The following {len(validationerrors["{0}".format(b)].index)}' \
                                            f' rows are missing entries for "{b}"; please correct this and resubmit,' \
                                            f' or press "pass" and note why these rows are being submitted with ' \
                                              f'blank entries for "{b}"'
            rowcounts.append(len(validationerrors["{0}".format(b)].index))
            valdf["{0}".format(b)] = validationerrors["{0}".format(b)]
            coldf["{0}".format(b)] = f'{b}'
            subtitle["{0}".format(b)] = "Row Check"


    #BR-xxx Sub LOB Blanks replaced by BR0012

    # schema checks if blank
    schema1={ 'Sub LOB': {'nullable': False, 'type': ['string'], 'empty': False  }, }
    print("schema1")
    # Cerberus functions
    v=Validator(schema1)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Sub LOB']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)
    s1 = pd.DataFrame(vdata.loc[rownums])

    if(len(s1.index)==0 ):
        v1='Sub LOB clear'
        cleared.append(v1)
    else:
        if len(s1.index) == 1:
            valmessage['Sub LOB'] = f'The following row is missing an entry for "Sub LOB";' \
                                            f' please correct this and resubmit, or press "pass" and note why this' \
                                            f' row is being submitted without "Sub LOB".'
        else:
            valmessage["Sub LOB"] = f'The following {len(s1.index)}' \
                                    f' rows are missing entries for "Sub LOB"; please correct this and resubmit,' \
                                    f' or press "pass" and note why these rows are being submitted without "Sub LOB".'
        rowcounts.append(len(s1.index))
        valdf["Sub LOB"] = s1
        coldf["Sub LOB"] = "Sub LOB"
        subtitle["Sub LOB"] = "Row Check"

 #   BR-xxx	B-Partner Blanks Blanks replaced by BR0012
  #  schema checks if blank
    schema2={ 'Business Partner Name': {'nullable': False,'type': 'string', 'empty': False}, }

    # Cerberus functions
    v=Validator(schema2)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Business Partner Name']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')
    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)
    s2 = pd.DataFrame(vdata.loc[rownums])

    if(len(s2.index)==0 ):
        v2='Business Partner Name'
        cleared.append(v2)
    else:
        if len(s2.index) == 1:
            valmessage['Business Partner Name'] = f'The following row is missing an entry for "Business Partner Name";'\
                                                f' please correct this and resubmit, or press "pass" and note why this'\
                                                f' row is being submitted without "Business Partner Name".'
        else:
            valmessage["Business Partner Name"] = f'The following {len(s2.index)}' \
                                          f' rows are missing entries for "Business Partner Name"; please correct this'\
                                          f' and resubmit or press "pass" and note why these rows are being submitted '\
                                          f'without "Business Partner Name".'
        rowcounts.append(len(s2.index))
        valdf["Business Partner Name"] = s2
        coldf["Business Partner Name"] = "Business Partner Name"
        subtitle["Business Partner Name"] = "Row Check"


   # BR-xxx	Product Provided Blanks replaced by BR0012, maybe edit into a concatenation of all product fields
    # schema checks if blank
    schema3 = { 'Product Name': {'nullable': False,'type': 'string', 'empty': False}, }

    # Cerberus functions
    v = Validator(schema3)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Product Name']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)
    s3 = pd.DataFrame(vdata.loc[rownums])

    if(len(s3.index)==0 ):
        v3='Product Name'
        cleared.append(v3)
    else:
        if len(s3.index) == 1:
            valmessage['Product Name'] = f'The following row is missing an entry for "Product Name",' \
                                            f' please correct this and resubmit; or press "pass" and note why this'\
                                            f' row is being submitted without product information.'
        else:
            valmessage["Product Name"] = f'The following {len(s3.index)}' \
                            f' rows are missing entries for "Product Name"; please correct this and resubmit,'\
                            f' or press "pass" and note why these rows are being submitted without product information.'
        rowcounts.append(len(s3.index))
        valdf["Product Name"] = s3
        coldf["Product Name"] = "Product Name"
        subtitle["Product Name"] = "Row Check"

    #BR016	Units of Risk Check (Earned or Written)
    schema4 = { 'Units of Risk (Earned)': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False, } }
    print("schema4")


    # Cerberus functions
    v = Validator(schema4)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Units of Risk (Earned)']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)
    s4 = pd.DataFrame(vdata.loc[rownums])

    if(len(s4.index)==0 ):
        v4='Units of Risk (Earned)'
        cleared.append(v4)
    else:
        if len(s4.index) == 1:
            valmessage['Units of Risk (Earned)'] = f'The following row is missing an entry for "Units of Risk ' \
                            f'(Earned)" please correct this and resubmit, or press "pass" and note why this' \
                            f' row is being submitted without "Units of Risk (Earned)".'
        else:
            valmessage['Units of Risk (Earned)'] = f'The following {len(s4.index)} rows are missing entries for ' \
                            f'"Units of Risk (Earned)", please correct this and resubmit, or press "pass" and note ' \
                            f'why these rows are being submitted without "Units of Risk (Earned)".'
        rowcounts.append(len(s4.index))
        valdf["Units of Risk (Earned)"] = s4
        coldf["Units of Risk (Earned)"] = "Units of Risk (Earned)"
        subtitle["Units of Risk (Earned)"] = "Row Check"

    schema4_2 = { 'Units of Risk (Written)': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False, } }
    print("schema42")
    # Cerberus functions
    v = Validator(schema4_2)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Units of Risk (Written)']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)
    s4_2 = pd.DataFrame(vdata.loc[rownums])

    if(len(s4_2.index)==0 ):
        v4_2='Units of Risk (Written)'
        cleared.append(v4_2)
    else:
        if len(s4_2.index) == 1:
            valmessage['Units of Risk (Written)'] = f'The following row is missing an entry for "Units of Risk '\
                                f'(Written)", please correct this and resubmit, or press "pass" and note why this' \
                                f' row is being submitted without "Units of Risk (Written)".'
        else:
            valmessage["Units of Risk (Written)"] = f'The following {len(s4_2.index)} rows are missing entries for ' \
                            f'"Units of Risk (Written)", please correct this and resubmit, or press "pass" and note' \
                            f' why these rows are being submitted without "Units of Risk (Written)".'

        rowcounts.append(len(s4_2.index))
        valdf["Units of Risk (Written)"] = s4_2
        coldf["Units of Risk (Written)"] = "Units of Risk (Written)"
        subtitle["Units of Risk (Written)"] = "Row Check"


    #BR016	Number of Policies (Earned)
    # schema checks if greater than 0
    schema5 = {'Number of Policies (Earned)': {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False,} }
    print("schema5")


    # Cerberus functions
    v = Validator(schema5)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Number of Policies (Earned)']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    s5 = pd.DataFrame(vdata.loc[rownums])

    if(len(s5.index)==0 ):
        v5='Number of Policies (Earned)'
        cleared.append(v5)
    else:
        if len(s5.index) == 1:
            valmessage['Number of Policies (Earned)'] = f'The following row is missing an entry for "Number of' \
                            f' Policies (Earned)", please correct this and resubmit, or press "pass" and note why this'\
                            f' row is being submitted without "Number of Policies (Earned)".'
        else:
            valmessage["Number of Policies (Earned)"] = f'The following {len(s5.index)} rows are missing entries for '\
                            f'"Number of Policies (Earned)", please correct this and resubmit, or press "pass" and '\
                            f'note why these rows are being submitted without "Number of Policies (Earned)".'
        rowcounts.append(len(s5.index))
        valdf["Number of Policies (Earned)"] = s5
        coldf["Number of Policies (Earned)"] = "Number of Policies (Earned)"
        subtitle["Number of Policies (Earned)"] = "Row Check"


    #BR016	Number of Policies (Written)
    # schema checks if greater than 0
    schema5_2 = {'Number of Policies (Written)':
                     {'nullable': False, 'type': 'number', 'min': 0.0001, 'required': False,}}
    print("schema52")
    # Cerberus functions
    v = Validator(schema5_2)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Number of Policies (Written)']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    s5_2 = pd.DataFrame(vdata.loc[rownums])

    if(len(s5_2.index)==0 ):
        v5_2='Number of Policies (Written)'
        cleared.append(v5_2)
    else:
        if len(s5_2.index) == 1:
            valmessage['Number of Policies (Written)'] = f'The following row is missing an entry for "Number of' \
                            f' Policies (Written)", please correct this and resubmit, or press "pass" and note' \
                            f' why this row is being submitted without "Number of Policies (Written)".'
        else:
            valmessage["Number of Policies (Written)"] = f'The following {len(s5_2.index)} rows are missing entries' \
                            f' for "Number of Policies (Written)", please correct this and resubmit, or press "pass"' \
                            f' and note why these rows are being submitted without "Number of Policies (Written)".'

        rowcounts.append(len(s5_2.index))
        valdf["Number of Policies (Written)"] = s5_2
        coldf["Number of Policies (Written)"] = "Number of Policies (Written)"
        subtitle["Number of Policies (Written)"] = "Row Check"

    #BR020	Commission Ratio Check
    vdata['comsub'] = vdata['Commission Ratio'].multiply(1, fill_value=0).replace(np.nan, '', regex=True)

    # schema checks if greater than 100%
    schema6 = { 'comsub': {'nullable': True, 'min': -.00000001, 'max': 1, 'required': False, } }
    print("schema6")
    # Cerberus functions
    v = Validator(schema6)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'comsub']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

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
            valmessage['comsub'] = f'The following row has a "Commission Ratio" not between 0% and 100% of ' \
                                   f'"Earned Revenue"; please correct this and resubmit, or press "pass" and note why' \
                                   f' this row is being submitted with the "Commission Ratio" as presented.'
        else:
            valmessage["comsub"] = f'The following {len(s6.index)} rows have a "Commission Ratio" not between 0% and ' \
                                   f'100% of "Earned Revenue"; please correct this and resubmit, or press "pass" and ' \
                                   f'note why these rows are being submitted with the "Commission Ratio" as presented.'

        rowcounts.append(len(s6.index))
        valdf["comsub"] = s6
        coldf["comsub"] = "Commission Ratio"
        subtitle["comsub"] = "Row Check"


    vdata['expsub'] = vdata['Expense Ratio'].multiply(1, fill_value=0).replace(np.nan, '', regex=True)

    #BR019	Expense Ratio Check
    # schema checks if greater than 100%
    schema7 = { 'expsub': {'nullable': True, 'min': -0.000000001, 'max': 1, 'required': False, } }
    print("schema7")
    # Cerberus functions
    v = Validator(schema7)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'expsub']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

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
            valmessage['expsub'] = f'The following row has a "Expense Ratio" not between 0% and 100% of ' \
                                   f'"Earned Revenue"; please correct this and resubmit, or press "pass" and note ' \
                                   f'why this row is being submitted with the "Expense Ratio" as presented.'
        else:
            valmessage["expsub"] = f'The following {len(s7.index)} rows have a "Expense Ratio" not between 0% and 100%'\
                                   f' of "Earned Revenue"; please correct this and resubmit, or press "pass" and note '\
                                   f'why these rows are being submitted with the "Expense Ratio" as presented.'

        rowcounts.append(len(s7.index))
        valdf["expsub"] = s7
        coldf["expsub"] = "Expense Ratio"
        subtitle["expsub"] = "Row Check"

        # BR006	Date Ranges
        # schema checks if date is correct
    schema8 = {'Reporting Date From INT': {'type': ['number'], 'max': 20190401, 'min': 20190401, 'required': False, }}
    print("schema8")
    # Cerberus functions
    v = Validator(schema8)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy

    try:
        vdata['Reporting Date From INT'] = pd.to_datetime(vdata['Reporting Date From']).dt.strftime("%Y%m%d").astype(int)
    except ValueError:
        tk.messagebox.showwarning(title="No Comment", message=f"Please check the that the dates do not "
                                                              f"contain mixed formats and reload.", )
        pass
        #os.startfile(spath)
        #sys.exit()


    vvdata = vdata[["Country", 'Reporting Date From INT']]


    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')


    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    s8 = pd.DataFrame(vdata.loc[rownums])

    if (len(s8.index) == 0):
        v8 = 'Reporting Date From'
        cleared.append(v8)
    else:
        if len(s8.index) == 1:
            valmessage['Reporting Date From'] = f'"Reporting Date From" is incorrect on this row, the Reporting Period'\
                                f' should begin on October 1st, 2019. This is a critical validation, email Global ' \
                                f'Portfolio Monitoring for instructions if you cannot provide the requested date range.'
        else:
            valmessage["Reporting Date From"] = f'"Reporting Date From" is incorrect on {len(s8.index)} rows, the ' \
                        f'Reporting Period should end on October 1st, 2019. This is a critical validation, email ' \
                        f'Global Portfolio Monitoring for instructions if you cannot provide the requested date range.'

        rowcounts.append(len(s8.index))
        valdf["Reporting Date From"] = s8
        coldf["Reporting Date From"] = "Reporting Date From"
        subtitle["Reporting Date From"] = "Row Check"

    #BR006	Date Ranges
    # schema checks if date is correct
    schema9 = { 'Reporting Date To INT': {'type': ['number'] , 'max': 20200930 , 'min': 20200930 , 'required': False, }}
    print("schema9")
    # Cerberus functions
    v = Validator(schema9)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    try:
        vdata['Reporting Date To INT'] = pd.to_datetime(vdata['Reporting Date To']).dt.strftime("%Y%m%d").astype(int)
    except ValueError:
        tk.messagebox.showwarning(title="No Comment", message=f"Please check the that the dates do not "
                                                              f"contain mixed formats and reload.", )
        pass

    vvdata = vdata[["Country", 'Reporting Date To INT']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    s9 = pd.DataFrame(vdata.loc[rownums])

    if (len(s9.index) == 0):
        v9 = 'Reporting Date To'
        cleared.append(v9)

    else:
        if len(s8.index) == 1:
            valmessage[
                'Reporting Date From'] = f'"Reporting Date From" is incorrect on this row, the Reporting Period' \
                                f' should end on September 30th, 2020. This is a critical validation, email Global '\
                                f'Portfolio Monitoring for instructions if you cannot provide the requested date range.'
        else:
            valmessage["Reporting Date From"] = f'"Reporting Date From" is incorrect on {len(s8.index)} rows, the ' \
                        f'Reporting Period should end on September 30th, 2020. This is a critical validation, email ' \
                        f'Global Portfolio Monitoring for instructions if you cannot provide the requested date range.'
        rowcounts.append(len(s9.index))
        valdf["Reporting Date To"] = s9
        coldf["Reporting Date To"] = "Reporting Date To"
        subtitle["Reporting Date To"] = "Row Check"


#BR006	Date Ranges
#schema checks if date is in correct range
    try:
        vdata['Date of Analysis INT'] = pd.to_datetime(vdata['Date of Analysis']).dt.strftime("%Y%m%d").astype(int)
    except ValueError:
        tk.messagebox.showwarning(title="No Comment", message=f"Please check the that the dates do not "
                                                              f"contain mixed formats and reload.", )
        pass

    schema10 = { 'Date of Analysis INT': {'type': ['number'] , 'max': 20201201 , 'min': 20201031 , 'required': False, }}
    print("schema10")
    # Cerberus functions
    v = Validator(schema10)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Date of Analysis INT']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')
    # empty list to fill
    rownums = []

    # append a row to rownums with each record that fails validation
    for idx, record in enumerate(data_dict):
        if not v.validate(record):
            rownums.append(idx)

    s10 = pd.DataFrame(vdata.loc[rownums])

    if len(s10.index) == 0:
        v10 = 'Date of Analysis'
        cleared.append(v10)
    else:
        if len(s10.index) == 1:
            valmessage['Date of Analysis'] = f'"Date of Analysis" is incorrect on this row, this should be after '\
                                                 f'the reporting period and before tomorrow.'
        else:
            valmessage['Date of Analysis'] = f'"Date of Analysis" is incorrect on {len(s10.index)} '\
                                                 f'rows, this should be after the reporting period and before tomorrow.'
        rowcounts.append(len(s10.index))
        valdf['Date of Analysis'] = s10
        coldf['Date of Analysis'] = 'Date of Analysis'
        subtitle['Date of Analysis'] = "Row Check"

    #BR013	Unaggregated Attributes

    """
    create a field to evaluate
    """
    vdata['Selected Fields for Duplicates'] = vdata['Country'].str.strip().fillna("")+\
        vdata['Type of Business'].str.strip().fillna("")+vdata['Type of Account'].str.strip().fillna("")\
        +vdata['Distribution Type'].str.strip().fillna("")+vdata['LOB'].str.strip().fillna("")\
        +vdata['Distribution Channel'].str.strip().fillna("")+vdata['Sub LOB'].str.strip().fillna("")\
        +vdata['Business Partner Name'].str.strip().fillna("")\
        +vdata['Business Partner ID Number'].str.strip().fillna("")+vdata['Product Name'].str.strip().fillna("")\
        +vdata['Product ID Number'].str.strip().fillna("")+vdata['Product Family'].str.strip().fillna("")\
        +vdata['Standard Product'].str.strip().fillna("")

    vcounts = pd.DataFrame(vdata['Selected Fields for Duplicates'].value_counts()).\
        reset_index().rename(columns={'Selected Fields for Duplicates' : 'Duplicate Count'})

    vdata = vdata.merge(vcounts, how= 'left',  left_on='Selected Fields for Duplicates', right_on='index',
                        suffixes=(False, False))

    schema11 = {'Duplicate Count': {'type': ['number'], 'max': 1,  'required': False, }}
    print("schema11")
    # Cerberus functions
    v = Validator(schema11)
    v.allow_unknown = True
    v.require_all = True

    # create a dictionary from the dataset (dict is a set where each item has an index)
    # this is not necessary since it is also created earlier, but I think it keeps it logical and tidy
    vvdata = vdata[["Country", 'Duplicate Count']]
    # create a dictionary from the dataset (dict is a set where each item has an index)
    data_dict = vvdata.to_dict(orient='records')

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
        valmessage['Duplicate Count'] = f'These combinations of Country, Distribution Channel, Sub LOB, B-Partner,' \
                                f' and Product are repeated, please aggregate and resubmit, or pass and notate.'
        rowcounts.append(len(s11.index))
        valdf['Duplicate Count'] = s11
        coldf['Duplicate Count'] = 'Duplicate Count'
        subtitle['Duplicate Count'] = "Attribute Check"

""" end validations section """

def validate(self):
    """Collects new lists from the data load."""
    if len(llock) == 0:

        global vmessagelist, sublist, vdflist, sdata, vdata, collist, sublist, clearedlist, rownlist, subtitle

        fetch_sdata()
        fetch_headers()

        vdata.append(sdata.replace(np.nan, '', regex=True))

        valid(vdata[-1], manfields, sdata)

        clearedlist = []
        clearedlist.append(cleared)

        vmessagelist = []
        v1_list = valmessage.values()
        vmessagelist = list(v1_list)

        vdflist = []
        v2_list = valdf.values()
        vdflist = list(v2_list)

        collist = []
        v3_list = coldf.values()
        collist = list(v3_list)

        sublist = []
        v4_list = subtitle.values()
        sublist = list(v4_list)

        rownlist = []
        rownlist.append(rowcounts)

    else:
        lock()

def change_df(self, input_val):
    """ This changes the DataFrame displayed in the UI pandas table.
    *this runs from buttons as a lambda*

    :param self: this is the table that needs to be changed
    :param input_val: a DataFrame, the final dataframe is set in teh below if statement
    :return: vdata[-1] in the pandastable
    :functions updateModel, TableModel: These pandastable functions change the DataFrame of the table
    :functions autoResizeColumns(): replaces the previous pandastable in the UI with the one created to memory in the line above

    'if len(llock) == 0:'

    Check if locked, this prevents it from running when the user has not explicitly
    indicated they want to reload the data
    """
    # Responds to button
    if len(llock) == 0:
        global vdata, sdata, ui_df

        vdata = []
        vdata.append(sdata.replace(np.nan, '', regex=True))
        ui_df = vdata[-1]
        self.updateModel(TableModel(vdata[-1]))
        self.showAll()
        self.clearSelected()
        self.autoResizeColumns()
    else:
        lock()


def submitbuttonaction(df):
    """
    see change_df() notes, this completes the same task but is callable from a mouse bound event (see python mouse
    documentation)
    :param df: current df to be changed
    :return: initial table load or refresh
    """
    if len(llock) == 0:
        global  vdata, ui_df, sdata, vdflist, table_df, uiv, vmessagelist, sublist, v1_list, tablemessage,\
            submessage, labelmessage2 , comlist, vn, rui, rdf, rmessage, rm, labelmessage3, rn

        vn = 0
        rn = 2
        vdata = []
        comlist = []

        fetch_sdata()
        fetch_headers()
        check_sheaders()
        make_reports(df)

        try:
            vdata.append(sdata.replace(np.nan, '', regex=True))
        except NameError:
            print("make an error message pop up, bad file or sheet")
        except AttributeError:
            pass
        ui_df = vdata[-1]
        ui.updateModel(TableModel(vdata[-1]))
        ui.showAll()
        ui.clearSelected()
        ui.autoResizeColumns()

        global controller, savecomment, vadd, headererrors
        check_sheaders()
        try:
            validate(df)
            table_df = vdflist[vn]
            uiv.updateModel(TableModel(vdflist[vn]))
            uiv.columncolors[collist[vn]] = '#FFFF00'
        except KeyError:
            table_df = headererrors
            uiv.updateModel(TableModel(headererrors))

        uiv.showAll()
        uiv.clearSelected()
        uiv.autoResizeColumns()

        try:
            disp_reps = {k: v for k, v in reports_dict.items()  if "B Partner & Product Combo" not in k }

            rdf = iter(disp_reps.values())
            rpos_df = next(rdf)
            rui.updateModel(TableModel(rpos_df))
            rm = iter(reports_dict.keys())
            rmessage = next(rm)
            labelmessage3.config(text=f"{rmessage} Report View")
        except KeyError:
            table_df = headererrors
            uiv.updateModel(TableModel(headererrors))

        rui.showAll()
        rui.clearSelected()
        rui.autoResizeColumns()

        try:
            labelmessage.config(text=f'{sublist[vn]} - {vmessagelist[vn]}')
            labelmessage2.config(text=f'{sublist[vn]} - {vmessagelist[vn]}')

        except NameError:
            pass
        lock()

    else:
        lock()

def resultsdraw():
    global cpos_df, collist, comlist , cui
    try:
        cpos_df = pd.DataFrame({"Validation Rule": collist, "Comments": comlist, "Row Counts": rowcounts, "Check Type":
            sublist, },)
    except NameError:
        cpos_df = pd.DataFrame({"Validation Rule": ["None"], "Comments": "All Clear!", "Row Counts": "None",
                                "Check Type": "None"},)

    cui.updateModel(TableModel(cpos_df ))
    cui.showAll()
    cui.clearSelected()
    cui.autoResizeColumns()



"""Button actions, used below in button option "command" , functions exist in scripts,
read functions as 'from [script in this folder] import [DataFrame]'"""


def subfolderopen():
    """System commands to open the folder location. Written to read the
    current file path so that it can work from any location where the folder tree is in tact.
    Note that the path has different logic for script vs exe execution.
    """
    os.startfile(spath)

def resourcefolderbuttonaction():
    """System commands to open the folder location. Written to read the
    current file path so that it can work from any location where the folder tree is in tact.
    Note that the path has different logic for script vs exe execution.
    """
    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Resources"))
    else:
        path = os.path.join(os.path.dirname(__file__), "Resources")
    os.startfile(path)

def dltemplatebuttonaction():
    """System commands to open the Template. Written to read the current file
    path so that it can work from any location where the template exists and the folder tree is in tact.
    """
    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Template"))
    else:
        path = os.path.join(os.path.dirname(__file__), "Template")

    file = os.listdir(path)

    filepath = os.path.join(str(path), str(file[0]))

    os.startfile(filepath)


def openreportbuttonaction():
    """System commands to open the Template. Written to read the current file
    path so that it can work from any location where the template exists and the folder tree is in tact.
    """
    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Report"))
    else:
        path = os.path.join(os.path.dirname(__file__), "Report")

    file = os.listdir(path)

    filepath = os.path.join(str(path), str(file[0]))

    os.startfile(filepath)



def emailer(text, subject, recipient):
    """Opens a prepopulated Outlook email if Outlook is open.

    Examples of how to use this are available here:

    https://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send

    Args:
        text (str): Body of email.

        subject (str): Subject of email.

        recipient (str): List of recipients example "<person.1@company.com>; <person.2@company.com>"
    *known issue: This will not run in Pycharm in Adminisrator Mode*
   """
    try:
        outlook = win32.GetActiveObject('Outlook.Application')
    except:
        outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Display(True)


def askassistbuttonaction():
    """callable from the tkinter button command, runs "emailer" function with correct variable values for the
    "ask for help" email

    *see uploadassistant.main.emailer(text, subject, recipient)*
    """

    emailer("", "Upload Application Assistance Required", "<Dana.Mark@allianz.com>; <angela.chenxx@allianz.com>; "
                                                "<Federico.Guerreschi@allianz.com>; <gavin.harmon@allianz.com>")


def finishbuttonaction():
    """Creates a zip file from the Output folder, opens the report ‘Global Portfolio Monitoring Report Views.xlsx’.
    The final submission email, and exits the application."""
    with ZipFile('Output.zip', 'w') as zipObj:
        # Iterate over all the files in directory
        for folderName, subfolders, filenames in os.walk(attachpath):
            for filename in filenames:
                # create complete filepath of file in directory
                filePath = os.path.join(folderName, filename)
                zipObj.write(filePath, basename(filename))
    try:
        os.rename(distzip, attachzip)
    except FileExistsError:
        pass
    except FileNotFoundError:
        pass

    emailer(f"Hello GPM,<br><br>I have completed the Data Collection using the Upload Assist"
            f"ant.<br><br>--Replace this text with any comments or explanations not captured in comments or survey.<br>"
            f"If you have made any currency conversions please explain with the LC to Euro figure.--"
            , "GPM 2020.Q3 Data Submission", " <Dana.Mark@allianz.com>;"
            " <angela.chenxx@allianz.com>; <Federico.Guerreschi@allianz.com>; <gavin.harmon@allianz.com>")

# mouse moves help control responses to user actions actions slight movement to trigger a command
def mousemove(self):
    """
    This is used to trigger as action in sequence immediately after a new page has been selected an loaded
    user effect - page navigated away from, waiting for new page to load, as opposed to "push a button and wait for
    something to happen"
    :param self: the button it is called from
    :return: inperceptable mouse movement
    """
    mouse.move(0, 1, absolute=False, duration=0)

# mouse moves help control responses to user actions actions slight movement to trigger a command
def mouseclick(self):
    """
    This is used to trigger as action insequence immediately after a new page has been selected an loaded
    user effect - page navigated away from, waiting for new page to load, as opposed to "push a button and wait for
    something to happen"
    :param self: the button it is called from
    :return: inperceptable mouse click
    """
    mouse.click()

# for is a command needs to run multiple functions *verify this is being used
def combine_funcs(*funcs):
    """This allows for multiple functions to be called in the command option within tkinter buttons"""

    def combined_func(*args, **kwargs):
        for f in funcs:
            f(*args, **kwargs)
    return combined_func


def resetvn():
    global vn
    vn = 0


def exportdata():
    """
    This runs when the user clicks "Finalize Submission" It does export all the data in the application as json files,
    but it also creates the report 'Global Portfolio Monitoring Report Views.xlsx'.
    """
    global reports_dict, hiddenbutton, rep_cols
    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Output"))
    else:
        path = os.path.join(os.path.dirname(__file__), "Output")

    file = f'us_orig{datetime.now().strftime("%m%d%Y%H%M%S")}.json'
    filepath = os.path.join(str(path), str(file))
    sdata.to_json(filepath, orient= "table")

    file = f'us_dat{datetime.now().strftime("%m%d%Y%H%M%S")}.json'
    filepath = os.path.join(str(path), str(file))
    vdata[-1].to_json(filepath, orient= "table")

    file = f'us_vcomments{datetime.now().strftime("%m%d%Y%H%M%S")}.json'
    filepath = os.path.join(str(path), str(file))
    cpos_df.to_json(filepath, orient="table")

    if hasattr(sys, 'frozen'):
        rpath = os.path.dirname(os.path.realpath(sys.executable).replace("dist", "Report"))
    else:
        rpath = os.path.join(os.path.dirname(__file__), "Report")

    file =  'Global Portfolio Monitoring Report Views.xlsx'

    filename = os.path.join(str(rpath),str( file))

    # generate nicely formatted excel file
    try:
        notes = {"Notes": pd.DataFrame({"Notes": ["Global Portfolio Monitoring Report Views. "
                          "", "", "",
                          "This report contains all standard views of the current data submission that will be"
                          " shared as part of Global Portfolio Monitoring's dashboards and reporting suite.", "",
                          "Please review the information on each sheet tab for accuracy and completeness. You are"
                          " encouraged to share this workbook locally for purposes of review.","",
                          "These reports may be shared with any Allianz employee that is authorized to review "
                          "Underwriting, Product, and profitability data.", "", "If there is any information that is"
                          " not accurate and sharable with executives and/or local managers and analysts, please take "
                          "this opportunity to correct the data submission before sending it.", "", "", "",
                          "Purpose:", "        The reports contained in this workbook are for you, the local Business "
                          "Unit representative's review and record. If you are confident in the reporting contained "
                          "on the sheet tabs, please save this report for later reference.", "", "        Only then "
                          "should you send Output.zip to GPM as an outlook email attachment.","" ,"        If there is "
                          "ever a need to correct or add context to GPM published materials you will have this as a "
                          "record of what was provided.", "", "", "", "Note:","         -Currency figures may have "
                          "formatting removed, no conversion operations have been performed between the template and "
                          "these reports.", "", "         -GPM may share this data as other views or transformed/aggregated within "
                          "the applicable Allianz Partners policies, procedures, and guidelines; always at the "
                          "direction of Global Product & Underwriting leadership.", "", "If you have questions please "
                          "email GPM at <dana.mark@allianz.com>, <angela.chen@allianz.com>, "
                          "<federico.guerreschi@allianz.com>, and/or <gavin.harmon@allianz.com>", "", "", "",
                          "Gavin Harmon 02-November-2020"]})}

        notes.update(reports_dict)
        reports_dict = notes

        nl.to_excel_ms(dfs=reports_dict, filename=filename)
        hiddenbutton.invoke()
    except PermissionError:
        tk.messagebox.showerror(title="Error", message="Please close any Excel or Outlook\n"
                                                       "files previously created by Upload\nAssistant to continue.")

class root(tk.Tk):
    """this is the main display, it is replaced by other pages as buttons get pushed or lifted to the user display
    The code for changing pages was derived from: http://stackoverflow.com/questions/7546050/switch-between-two-frames
    -in-tkinter
    Tkinter basics can be found here https://docs.python.org/3/library/tk.html  """

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        # app title that displays in the window
        self.winfo_toplevel().title("Upload Assistant")

        # all pages must exist here

        for F in (p01StartPage, p02LoadPage, p04DataSetViewer, p05ReportViewer, p06ValidationView, p07commentpage, p08ValidationReport, p09SaveComments ,
                  p10SurveyOne , p11SurveyTwo, p12SurveyThree, p13SurveyFour, p14SurveyFive, p15ExitPage ):

            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(p01StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

def launchinsts():
    """Launches image viewer for the "Instructions" slides."""
    global my_label
    global button_forward
    global button_exit
    global button_back

    newWindow = tk.Toplevel(app)
    newWindow.title("New Window")

    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", ""))
    else:
        path = os.path.dirname(__file__)

    bgimage0 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions1.png")))
    bgimage1 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions2.png")))
    bgimage2 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions3.png")))
    bgimage3 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions4.png")))
    bgimage4 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions5.png")))
    bgimage5 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions6.png")))
    bgimage6 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/instructions7.png")))

    image_list = [bgimage0, bgimage1, bgimage2, bgimage3, bgimage4, bgimage5, bgimage6, ]

    my_label = tk.Label(newWindow, image=bgimage0)
    my_label.pack()

    def forward(image_number):
        global my_label
        global button_forward
        global button_exit
        global button_back

        my_label.pack_forget()
        button_forward.place_forget()
        button_exit.place_forget()
        button_back.place_forget()
        my_label = tk.Label(newWindow, image=image_list[image_number-1])
        button_forward = tk.Button(newWindow, text="Next",  font=('Helvetica', 15), bg='#004a93', fg="white",
                                   command=lambda:  forward(image_number+1))
        button_forward.place(relx= .61, rely=.9, relwidth=0.075, )
        button_exit = tk.Button(newWindow, text="Exit", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command= newWindow.destroy)
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command=lambda:  back(image_number-1))
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

        if image_number == len(image_list):
            button_forward = tk.Button(newWindow, font=('Helvetica', 15), bg='#004a93', fg="white",
                                       text="Next", state=DISABLED)
            button_forward.place(relx=.61, rely=.9, relwidth=0.075, )

        my_label.pack()
        button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    def back(image_number):
        global my_label
        global button_forward
        global button_exit
        global button_back

        my_label.pack_forget()
        button_forward.place_forget()
        button_exit.place_forget()
        button_back.place_forget()
        my_label = tk.Label(newWindow, image=image_list[image_number-1])
        button_forward = tk.Button(newWindow, text="Next", font=('Helvetica', 15), bg='#004a93', fg="white",
                                   command=lambda:  forward(image_number+1))
        button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
        button_exit = tk.Button(newWindow, text="Exit", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command= newWindow.destroy )
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command=lambda:  back(image_number-1))
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

        if image_number == 1:
            button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                                    state=DISABLED)

        my_label.pack()
        button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    button_forward = tk.Button(newWindow, text="Next", font=('Helvetica', 15), bg='#004a93', fg="white",
                               command=lambda: forward(2))
    button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
    button_exit = tk.Button(newWindow, text="Exit", font=('Helvetica', 15), bg='#004a93', fg="white",
                            command=newWindow.destroy)
    button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
    button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                            state=DISABLED)
    button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
    button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
    button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    app.mainloop()


def launchdefs():
    """Launches image viewer for the "Metric Definitions" slides."""
    global my_label
    global button_forward
    global button_exit
    global button_back

    newWindow = tk.Toplevel(app)
    newWindow.title("New Window")

    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", ""))
    else:
        path = os.path.dirname(__file__)

    met_defs0 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs1.png")))
    met_defs1 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs2.png")))
    met_defs2 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs3.png")))
    met_defs3 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs4.png")))
    met_defs4 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs5.png")))
    met_defs5 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs6.png")))
    met_defs6 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs7.png")))
    met_defs7 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs8.png")))
    met_defs8 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs9.png")))
    met_defs9 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs10.png")))
    met_defs10 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs11.png")))
    met_defs11 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs12.png")))
    met_defs12 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs13.png")))
    met_defs13 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs14.png")))
    met_defs14 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs15.png")))
    met_defs15 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/met_defs16.png")))


    image_list = [met_defs0, met_defs1, met_defs2, met_defs3, met_defs4,met_defs5, met_defs6,
                  met_defs7, met_defs8, met_defs9,met_defs10, met_defs11, met_defs12, met_defs13,
                  met_defs14,met_defs15, ]

    my_label = tk.Label(newWindow, image=met_defs0)
    my_label.pack()

    def forward(image_number):
        global my_label
        global button_forward
        global button_exit
        global button_back

        my_label.pack_forget()
        button_forward.place_forget()
        button_exit.place_forget()
        button_back.place_forget()
        my_label = tk.Label(newWindow, image=image_list[image_number-1])
        button_forward = tk.Button(newWindow, text="Next",  font=('Helvetica', 15), bg='#004a93', fg="white",
                                   command=lambda:  forward(image_number+1))
        button_forward.place(relx= .61, rely=.9, relwidth=0.075, )
        button_exit = tk.Button(newWindow, text="Exit", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command= newWindow.destroy)
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command=lambda:  back(image_number-1))
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

        if image_number == len(image_list):
            button_forward = tk.Button(newWindow, font=('Helvetica', 15), bg='#004a93', fg="white",
                                       text="Next", state=DISABLED)
            button_forward.place(relx=.61, rely=.9, relwidth=0.075, )

        my_label.pack()
        button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    def back(image_number):
        global my_label
        global button_forward
        global button_exit
        global button_back

        my_label.pack_forget()
        button_forward.place_forget()
        button_exit.place_forget()
        button_back.place_forget()
        my_label = tk.Label(newWindow, image=image_list[image_number-1])
        button_forward = tk.Button(newWindow, text="Next", font=('Helvetica', 15), bg='#004a93', fg="white",
                                   command=lambda:  forward(image_number+1))
        button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
        button_exit = tk.Button(newWindow, text="Exit", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command= newWindow.destroy )
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                                command=lambda:  back(image_number-1))
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

        if image_number == 1:
            button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                                    state=DISABLED)

        my_label.pack()
        button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
        button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
        button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    button_forward = tk.Button(newWindow, text="Next", font=('Helvetica', 15), bg='#004a93', fg="white",
                               command=lambda: forward(2))
    button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
    button_exit = tk.Button(newWindow, text="Exit", font=('Helvetica', 15), bg='#004a93', fg="white",
                            command=newWindow.destroy)
    button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
    button_back = tk.Button(newWindow, text="Back", font=('Helvetica', 15), bg='#004a93', fg="white",
                            state=DISABLED)
    button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    button_forward.place(relx=.61, rely=.9, relwidth=0.075, )
    button_exit.place(relx=.45, rely=.9, relwidth=0.075, )
    button_back.place(relx=.28, rely=.9, relwidth=0.075, )

    app.mainloop()


def launchabout():
    """Launches image viewer for the "About GPM" slide."""
    global my_label
    global button_forward
    global button_exit
    global button_back

    newWindow = tk.Toplevel(app)
    newWindow.title("New Window")

    if hasattr(sys, 'frozen'):
        path = os.path.dirname(os.path.realpath(sys.executable).replace("dist", ""))
    else:
        path = os.path.dirname(__file__)

    about_defs0 = ImageTk.PhotoImage(Image.open(os.path.join(path, "Images/mission.png")))

    my_label = tk.Label(newWindow, image=about_defs0)
    my_label.pack()

    app.mainloop()


class p01StartPage(tk.Frame):
    """
    This is the start page that launches first.
    """

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        HEIGHT= 600
        WIDTH =1400

        home = tk.Canvas(self, height=HEIGHT, width=WIDTH, bd=0, highlightthickness=0 )
        home.pack(fill="both", expand=True, anchor="s")
        home.bind("<Enter>", unlock)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(home, image=background_image)
        background_label.image = background_image
        background_label.place( relwidth=1, relheight=1)

        label = tk.Label(home, text="GPM 2020.Q3 Upload Assistant", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        instructionsbutton = tk.Button(home, text="Submission\nInstructions", font=('Helvetica', 16), bg='#004a93',
                                       fg="white",  command= launchinsts)
        instructionsbutton.place(relx=0.135, rely=.295, relwidth=0.14, relheight=0.07)

        MetricDefsbutton = tk.Button(home, text="Metric\nDefinitions", font=('Helvetica', 16), bg='#004a93', fg="white",
                                     command= launchdefs )
        MetricDefsbutton.place(relx=0.415, rely=.295, relwidth=0.14, relheight=0.07)

        askassistbutton = tk.Button(home, text="Ask for\nAssistance", font=('Helvetica', 16), bg='#004a93', fg="white",
                                    command=askassistbuttonaction)
        askassistbutton.place(relx=0.695, rely=.295, relwidth=0.14, relheight=0.07)

        dltemplatebutton = tk.Button(home, text="Open\nTemplate", font=('Helvetica', 16), bg='#cbe1f7', fg="black",
                                     command=dltemplatebuttonaction)
        dltemplatebutton.place(relx=0.135, rely=.48, relwidth=0.14, relheight=0.07)

        subfolderbutton = tk.Button(home, text="Open Submission\nLocation", font=('Helvetica', 14), bg='#cbe1f7',
                                    fg="black", command=subfolderopen)
        subfolderbutton.place(relx=0.415, rely=.48, relwidth=0.14, relheight=0.07)

        resourcesbutton = tk.Button(home, text="Resources", font=('Helvetica', 16), bg='#cbe1f7', fg="black",
                                    command=resourcefolderbuttonaction)
        resourcesbutton.place(relx=0.695, rely=.48, relwidth=0.14, relheight=0.07)

        about_gpm_button = tk.Button(home, text="About Global\nPortfolio Monitoring", font=('Helvetica', 16),bg= '#006094', fg="white",
                                    command=launchabout)
        about_gpm_button.place(relx=0.7, rely=.8, relwidth=0.16, relheight=0.12)

        submitbutton = tk.Button(home, text="Load Submission", font=('Helvetica', 22), bg='#ff7300', fg="black",
                                 command=  lambda: controller.show_frame(p02LoadPage))
        submitbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07)


class p02LoadPage(tk.Frame):
    """
    This is a copy of the start page with a "loading data" button. This informs the user that the "load submission"
    button has been pressed and they are waiting for the application to run.
    """

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        HEIGHT= 600
        WIDTH =1400

        def autopage(self):
            submitbutton.invoke()

        home = tk.Canvas(self, height=HEIGHT, width=WIDTH, bd=0, highlightthickness=0 )
        home.pack(fill="both", expand=True, anchor="s")
        home.bind("<Enter>", unlock)
        home.bind("<Enter>", autopage)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(home, image=background_image)
        background_label.image = background_image
        background_label.place( relwidth=1, relheight=1)

        label = tk.Label(home, text="GPM 2020.Q3 Upload Assistant", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        instructionsbutton = tk.Button(home, text="Submission\nInstructions", font=('Helvetica', 16), bg='#004a93',
                                       fg="white")#, command=lambda: controller.show_frame(p03MetricDefinitions))
        instructionsbutton.place(relx=0.135, rely=.295, relwidth=0.14, relheight=0.07)

        MetricDefsbutton = tk.Button(home, text="Metric\nDefinitions", font=('Helvetica', 16), bg='#004a93', fg="white",
                                     command=lambda: controller.show_frame(p04DataSetViewer))
        MetricDefsbutton.place(relx=0.415, rely=.295, relwidth=0.14, relheight=0.07)

        askassistbutton = tk.Button(home, text="Ask for\nAssistance", font=('Helvetica', 16), bg='#004a93', fg="white",
                                    command=askassistbuttonaction)
        askassistbutton.place(relx=0.695, rely=.295, relwidth=0.14, relheight=0.07)

        dltemplatebutton = tk.Button(home, text="Open\nTemplate", font=('Helvetica', 16), bg='#cbe1f7', fg="black",
                                     command=dltemplatebuttonaction)
        dltemplatebutton.place(relx=0.135, rely=.48, relwidth=0.14, relheight=0.07)

        subfolderbutton = tk.Button(home, text="Open Submission\nLocation", font=('Helvetica', 14), bg='#cbe1f7',
                                    fg="black", command=subfolderopen)
        subfolderbutton.place(relx=0.415, rely=.48, relwidth=0.14, relheight=0.07)

        resourcesbutton = tk.Button(home, text="Resources", font=('Helvetica', 16), bg='#cbe1f7', fg="black",
                                    command=resourcefolderbuttonaction)
        resourcesbutton.place(relx=0.695, rely=.48, relwidth=0.14, relheight=0.07)

        about_gpm_button = tk.Button(home, text="About Global\nPortfolio Monitoring", font=('Helvetica', 16),bg= '#006094', fg="white",
                                    command=launchabout)
        about_gpm_button.place(relx=0.7, rely=.8, relwidth=0.16, relheight=0.12)

        df = pd.DataFrame( {'Status': [f'No data to display.', 'Please check that the Excel file is saved in the',
                                    f'{spath} folder and that the submission data is on a',
                                   f'sheet tab named "Ptf_Monitoring_GROSS_Reins" (No spaces).', ]})

        submitbutton = tk.Button(home, text="Data loading...", font=('Helvetica', 22), bg='#D3D3D3', fg="black",
                    command= combine_funcs(lambda:submitbuttonaction(self), lambda: controller.show_frame(p04DataSetViewer)))
        submitbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07)
        submitbutton.bind("<Button-1>", submitbuttonaction)

#
# class p03MetricDefinitions(tk.Frame):
#     """
#     Metric Definitions Page
#     """
#
#     def __init__(self, parent, controller):
#         tk.Frame.__init__(self, parent)
#         metricdefs = tk.Frame(self, height=850, width=1150, bg='#dedede')
#         metricdefs.pack(fill="both", expand=True)
#
#         background_image = tk.PhotoImage(file=bgimage)
#         background_image.image = 31
#         background_label = tk.Label(metricdefs, image=background_image, height=900, width=1400, )
#         background_label.image = background_image
#         background_label.place(relwidth=1, relheight=1)
#
#
#         label = tk.Label(metricdefs, text="Metric Definitions", font=('Helvetica', 22), bg='#004a93', fg="white", )
#         label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)
#
#         homebutton = tk.Button(metricdefs, text="Start Over", font=('Helvetica', 16), bg='#004a93', fg="white",
#                                command=lambda: controller.show_frame(p01StartPage))
#         homebutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )


class p04DataSetViewer(tk.Frame):
    """
    Page that displays the full data set data table.
    """

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150,  )
        submit.pack(fill="both", expand=True)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Submission Results", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.06)

        box = tk.Frame(submit, bg='#004a93', )
        box.place(relx=0.16, rely=0.2675, relwidth=0.65, relheight=0.55, )

        class UserInterface(Table):

            def change_df(self, input_val):
               submitbuttonaction(self)

        global ui, ui_df, pos_df


        pos_data = {'Status': [f'No data to display.', 'Please check that the Excel file is saved in the',
                                f'{spath}', f'folder and that the submission data is on a',
                               f' sheet tab named "Ptf_Monitoring_GROSS_Reins" (No spaces).', ]}

        pos_df = pd.DataFrame(data=pos_data)

        ui_df = pos_df

        ui = UserInterface(box, dataframe=pos_df, showtoolbar=True, showstatusbar=True,)
        ui.show()

        nextbutton = tk.Button(submit, text="View Summary\nReports", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=lambda: controller.show_frame(p05ReportViewer))
        nextbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )
        submit.bind("<Enter>", mousemove)

        homebutton = tk.Button(submit, text="Reload Submission\n"
                                            "and Restart", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=lambda: controller.show_frame(p01StartPage))
        homebutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        homebutton.bind("<Button-1>", unlock)

        subfolderbutton = tk.Button(submit, text="Open Submission\n Folder", font=('Helvetica', 16), bg='#004a93',
                                    fg="white",
                                    command=subfolderopen)
        subfolderbutton.place(relx=0.845, rely=.444, width=175, height=75, )

        instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93', fg="white",
                                   command= launchinsts)
        instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )


class p05ReportViewer(tk.Frame):
    """
    Page that displays the reports as a pandastable.
    """

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        global labelmessage3

        submit = tk.Frame(self, height=850, width=1150,)
        submit.pack(fill="both", expand=True)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        labelmessage3 = tk.Label(submit, text=f"Report View", font=('Helvetica', 22), bg='#004a93', fg="white", )
        labelmessage3.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.04)
        labelmessage4 = tk.Label(submit, text=f"Gross of Reinsurance view in Local Currency", font=('Helvetica', 14), bg='#004a93', fg="white", )
        labelmessage4.place(relx=0.26, rely=.1765, relwidth=0.475, relheight=0.04)


        box = tk.Frame(submit, bg='#004a93', )
        box.place(relx=0.16, rely=0.2675, relwidth=0.65, relheight=0.55, )

        class UserInterface(Table):

            def change_df(self, input_val):
               submitbuttonaction(self)

        global rui, rui_df, rpos_df, rdf, reports_dict, reports_dict_float, rn

        rn = 2


        rpos_data = {'Status': [f'No data to display.', 'Please check that the Excel file is saved in the',
                               f'{spath} folder and that the submission data is on a',
                               f' sheet tab named "Ptf_Monitoring_GROSS_Reins" (No spaces).', ]}

        rpos_df =pd.DataFrame(rpos_data)

        try:
            rdf = iter(reports_dict.values())
            rpos_df = next(rdf)
            rui_df = pd.DataFrame(data=rpos_df)

        except NameError:
            pass
        except StopIteration:
            print("aa")

        rui = Table(box, dataframe=rpos_df, showtoolbar=True, showstatusbar=True, )
        rui.show()

        def reportnext():
            global rposdf, reports_dict, reports_dict_float, rn, rm, rmessage, labelmessage3

            try:
                if rn == len(reports_dict.items())-1:
                    nextbutton3.config(text="Confirm Report and\nBegin Validation")
                else:
                    nextbutton3.config(text="Confirm as Correct")
                    pass
                rui.updateModel(TableModel(next(rdf)))
                rui.showAll()
                rui.clearSelected()
                rui.autoResizeColumns()
                rn = rn +1

            except StopIteration:
                hiddenbutton.invoke()

        def changerepmessage():
            try:
                rmessage = next(rm)
                labelmessage3.config(text=f"{rmessage} Report View")
            except StopIteration:
                pass


        nextbutton3 = tk.Button(submit, text="Next Report", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=combine_funcs(changerepmessage, reportnext))
        nextbutton3.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )

        submit.bind("<Enter>", mousemove)

        homebutton = tk.Button(submit, text="Reload Submission\n"
                                            "and Restart", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=lambda: controller.show_frame(p01StartPage))
        homebutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        homebutton.bind("<Button-1>", unlock)

        subfolderbutton = tk.Button(submit, text="Open Submission\n Folder", font=('Helvetica', 16), bg='#004a93',
                                    fg="white",
                                    command=subfolderopen)
        subfolderbutton.place(relx=0.845, rely=.444, width=175, height=75, )

        instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93', fg="white",
                                   command= launchinsts)
        instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )

        hiddenbutton = tk.Button(submit, text="Begin Validation", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command= lambda: controller.show_frame(p06ValidationView))
        hiddenbutton.place(relx=0, rely= 0, relwidth=0, relheight=0, )

class p06ValidationView(tk.Frame):
    """
    Show Validation result window, loop through.
    """
    global vdflist, vmessagelist, sublist, tabledf, uiv, tablemessage, vlock, sublist, submessage
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        def skiplast(self):
            """this forces forces an error at the coincidental flow point where the user has completed navigation and
            needs to move the the review results section.

            Troubleshooting hint:
            If this goes wrong you will either get a repeat validation 1 time
            at the end, or upon resubmission you will skip this page altogether"""
            global exitbutton
            try:
                if vmessagelist[vn] == vmessagelist[-1]:
                    pass
                else:
                    pass
            except IndexError:
                nextbutton.invoke()
            except NameError:
                nextbutton.config(command=header_warn)
            else:
                lock

        submit = tk.Frame(self, height=850, width=1150,  )
        submit.pack(fill="both", expand=True)
        submit.bind("<Enter>",skiplast)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Submission Results", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.06)

        box = tk.Frame(submit, bg='#004a93', )
        box.place(relx=0.16, rely=0.2875, relwidth=0.65, relheight=0.55)

        global vmessagelist,  tablemessage, labelmessage, sublist

        try:
            tablemessage = f'{sublist[vn]} - {vmessagelist[vn]}'
        except NameError:
            tablemessage = "Data Not Loaded"

        labelmessage = tk.Label(submit, text=tablemessage, font=('Helvetica', 12,), wraplength=700, justify='center',
                                bg='white', fg='black', anchor='center')
        labelmessage.place(relx=0.16, rely=0.205, relwidth=0.65, relheight=0.08, )

        global table_df, vpos_df, vui_df, uiv, vdflist, collist , homebutton, spath


        vpos_data = {'Status': [f'No data to display. Please check that the Excel file is saved in '
                                    f'the {spath} folder and that the submission data is on a sheet tab named'
                                f'"Ptf_Monitoring_GROSS_Reins" (No spaces).', ]}

        try:
            vpos_df = vdflist[vn]
        except NameError:
            vpos_df = pd.DataFrame(data=vpos_data)

        uiv = Table(box, dataframe= vpos_df, showtoolbar=True, showstatusbar=True)
        try:
            uiv.columncolors[collist[vn]] = '#FFFF00'
        except IndexError:
            pass
        except NameError:
            pass
        uiv.show()

        def header_warn():
            tk.messagebox.showwarning(title="No Comment", message=f"Please check the headers of latest file saved in"
                        f" {spath}. This error indicates a misnamed header.", )
            os.startfile(spath)
            sys.exit()

        nextbutton = tk.Button(submit, text="Pass with Comment", font=('Helvetica', 16), bg='#ff7300', fg="black",
                               command=lambda: controller.show_frame(p07commentpage))
        nextbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )

        homebutton = tk.Button(submit, text="Reload Submission\n"
                                            "and Restart", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=lambda: controller.show_frame(p01StartPage))
        homebutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        homebutton.bind("<Button-1>", unlock)

        subfolderbutton = tk.Button(submit, text="Open Submission\n Folder", font=('Helvetica', 16), bg='#004a93',
                                    fg="white",
                                    command=subfolderopen)
        subfolderbutton.place(relx=0.845, rely=.444, width=175, height=75, )

        instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93', fg="white",
                                   command= launchinsts)
        instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )



class p07commentpage(tk.Frame):
    """
    Comment page for a validation pass , loop through.
    """
    global vdflist, vmessagelist, tabledf, uiv, tablemessage, label, submessage, sublist

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        def fwdrefresh(self):
            global exitbutton
            if len(llock)==0:
                exitbutton.invoke()
            else:
                lock

        submit = tk.Frame(self, height=850, width=1150,  )
        submit.pack(fill="both", expand=True)
        submit.bind("<Enter>",fwdrefresh)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)


        global label
        label = tk.Label(submit, text="Data Submission Results", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.06)

        global labelmessage2, tablemessage

        labelmessage2 = tk.Label(submit, text=tablemessage, font=('Helvetica', 12,), wraplength=700, justify='center',
                                bg='white', fg='black', anchor='center')
        labelmessage2.place(relx=0.16, rely=0.205, relwidth=0.65, relheight=0.08, )


        box = tk.Frame(submit, bg='#004a93', )
        box.place(relx=0.16, rely=0.2875, relwidth=0.65, relheight=0.55)

        global vmessagelist, sublist, submessage,  comlist

        try:
            tablemessage = f'{sublist[vn]} - {vmessagelist[vn]}'
            submessage = sublist[vn]
        except NameError:
            tablemessage = "Data Not Loaded"
            submessage = "Data Not Loaded2"



        global commbox
        commbox = Entry(submit, width=300)
        commbox.place(relx=0.16, rely=0.45, relwidth=0.65, relheight=0.06, )
        commbox.insert(0, "Please replace this text with an explanation of why the validation cannot be cleared.")

        def clickfalse():
            comment = commbox.get()
            if comment == "Please replace this text with an explanation of why the validation cannot be cleared."\
                    or comment == "":
                tk.messagebox.showwarning(title="No Comment", message="Please edit the comment\ntext before moving on.", )
            else:
                falsebutton.invoke()

        def savecomment():
            comment = commbox.get()
            comlist.append(comment)
            commbox.delete(0, len(comment))
            commbox.insert(0, "Please replace this text with an explanation of why the validation cannot be cleared.")

        def vadd():
            global vn, vx, table_df, vdflist, uiv, labelmessage, labelmessage2, collist, master, label, commbox,\
                nextbutton, exitbutton, vlock, sublist

            vn = vn + 1
            vx = min(len(vdflist) - 1, vn)
            if vn > len(vdflist) - 1:
                vlock={1}
                label.config(text="Validations Complete")
                label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

                commbox.config(width=0)
                commbox.place(relwidth=0.000001, relheight=0.00000001, )
                nextbutton.place(relwidth=0.000001, relheight=0.00000001, )
                commbox.insert(0, "Please replace this text with an explanation of why the validation cannot be cleared.")

                nextbutton = tk.Button(submit, text="See Results",  bg='#ff7300' , font=('Helvetica', 22) , fg="black",
                                   command=combine_funcs(lambda: controller.show_frame(p08ValidationReport), resetvn, resultsdraw))
                nextbutton.place(relx=0.38, rely=.45, relwidth=0.2, relheight=0.125, )

                submissionfolderbutton = tk.Button(submit, text="Submission Folder", font=('Helvetica', 16), bg='#004a93', fg="white",
                                       command=subfolderopen)
                #submissionfolderbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )
                submissionfolderbutton.place(relx=0.395, rely=.85, relwidth=0, relheight=0, )

                exitbutton = tk.Button(submit, text="Reload Submission\n"
                                            "and Restart", font=('Helvetica', 16), bg='#004a93', fg="white",
                                       command= combine_funcs(lambda: controller.show_frame(p01StartPage), rebuild ))
                exitbutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
                exitbutton.bind("<Button-1>", unlock)

                subfolderbutton = tk.Button(submit, text="Open Submission\n Folder", font=('Helvetica', 16),
                                            bg='#004a93',
                                            fg="white",
                                            command=subfolderopen)
                subfolderbutton.place(relx=0.845, rely=.444, width=175, height=75, )

                instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93',
                                               fg="white",
                                               command= launchinsts)
                instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )

                labelmessage2.config(text=f'Click "See Results" to se a validation report with your comments that '
                                          f'you can copy and paste if you choose.')

                # labelmessage.config(text="")
                # labelmessage.place(relx=0, rely=0, relwidth=0.000001, relheight=0.00000001, )
                #
                # labelmessage2.config(text="")
                # labelmessage2.place(relx=0, rely=0, relwidth=0.000001, relheight=0.00000001, )

            else:
                table_df = vdflist[vx]
                uiv.updateModel(TableModel(vdflist[vx]))
                for idx, val in enumerate(collist):
                    uiv.columncolors[collist[idx]] = '#F4F4F3'

                uiv.columncolors[collist[vx]] = '#FFFF00'
                uiv.autoResizeColumns()

                labelmessage.config(text=f'{sublist[vx]} - {vmessagelist[vx]}')
                labelmessage2.config(text=f'{sublist[vx]} - {vmessagelist[vx]}')

        def rebuild():
            global vdflist, vmessagelist, sublist, tabledf, uiv, tablemessage, label, vn, commbox, nextbutton,\
                exitbutton, submessage, labelmessage2

            vn=0
            label = tk.Label(submit, text="Data Submission Results", font=('Helvetica', 22), bg='#004a93', fg="white")
            label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.06)

            box = tk.Frame(submit, bg='#004a93', )
            box.place(relx=0.16, rely=0.2875, relwidth=0.65, relheight=0.55)

            global vmessagelist, sublist, tablemessage, comlist, submessage, sublist, submessage, labelmessage2

            labelmessage2 = tk.Label(submit, text=tablemessage, font=('Helvetica', 12,), wraplength=700,
                                     justify='center',
                                     bg='white', fg='black', anchor='center')
            labelmessage2.place(relx=0.16, rely=0.205, relwidth=0.65, relheight=0.08, )

            try:
                tablemessage = f'{sublist[vn]} - {vmessagelist[vn]}'
                submessage = sublist[vn]
            except NameError:
                tablemessage = "Data Not Loaded"
                submessage = "Data Not Loaded3"


            global commbox
            commbox = Entry(submit, width=300)
            commbox.place(relx=0.16, rely=0.45, relwidth=0.65, relheight=0.06, )
            commbox.insert(0, "Please replace this text with an explanation of why the validation cannot be cleared.")

            nextbutton = tk.Button(submit, text="Save Comment\nand Continue", font=('Helvetica', 16), bg='#ff7300', fg="black",
                                   command=clickfalse)
            nextbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )

            falsebutton = tk.Button(submit, text="Save Comment\nand Continue", font=('Helvetica', 16), bg='#ff7300', fg="black",
                                   command=combine_funcs(savecomment, lambda: controller.show_frame(p06ValidationView),  vadd))
            falsebutton.place(relx=0, rely=0, relwidth=0, relheight=0, )

            exitbutton = tk.Button(submit, text="Reload Submission\n"
                                            "and Restart", font=('Helvetica', 16), bg='#004a93', fg="white",
                                   command= combine_funcs(lambda: controller.show_frame(p01StartPage), rebuild ))
            exitbutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )

            subfolderbutton = tk.Button(submit, text="Open Submission\n Folder", font=('Helvetica', 16), bg='#004a93',
                                        fg="white",
                                        command=subfolderopen)
            subfolderbutton.place(relx=0.845, rely=.444, width=175, height=75, )

            instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93',
                                           fg="white", command= launchinsts)
            instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )

        global table_df, vpos_df, vui_df, uiv, vdflist, collist, comlist, nextbutton, exitbutton

        nextbutton = tk.Button(submit, text="Save Comment\nand Continue", font=('Helvetica', 16), bg='#ff7300',
                               fg="black",
                               command=clickfalse)
        nextbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )

        falsebutton = tk.Button(submit, text="Save Comment\nand Continue", font=('Helvetica', 16), bg='#ff7300',
                                fg="black",
                                command=combine_funcs(savecomment, lambda: controller.show_frame(p06ValidationView), vadd))
        falsebutton.place(relx=0, rely=0, relwidth=0, relheight=0, )

        exitbutton = tk.Button(submit, text="Start Over", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=lambda: controller.show_frame(p07commentpage) )
        exitbutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        exitbutton.bind("<Button-1>", unlock)

        subfolderbutton = tk.Button(submit, text="Open Submission\n Folder", font=('Helvetica', 16), bg='#004a93',
                                    fg="white",
                                    command=subfolderopen)
        subfolderbutton.place(relx=0.845, rely=.444, width=175, height=75, )

        instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93', fg="white",
                                   command=  launchinsts)
        instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )


class p08ValidationReport(tk.Frame):
    """
    The review comments page.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Validation Results and Comments", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        box = tk.Frame(submit, bg='#004a93', )
        box.place(relx=0.16, rely=0.2675, relwidth=0.65, relheight=0.55)

        global cui, cui_df, cpos_df, collist, comlist

        try:
            cpos_df = pd.DataFrame({"Validation Rule": collist, "Comments": comlist, "Row Counts": rowcounts, "Check Type": sublist, }, )
        except NameError:
            cpos_df = pd.DataFrame({"Validation Rule": ["None"], "Comments": "All Clear!", "Row Counts": "None", "Check Type": "None" }, )

        cui = Table(box, dataframe=cpos_df, showtoolbar=True, showstatusbar=True)
        cui.show()

        finishvalbutton = tk.Button(submit, text="Finalize\nSubmission", font=('Helvetica', 16), bg='#ff7300',
                                    fg="black", command=  exportdata)
        finishvalbutton.place(relx=0.845, rely=.444, width=175, height=75, )

        # nextbutton = tk.Button(submit, text="View Reports", font=('Helvetica', 16), bg='#004a93', fg="white",
        #                        command=lambda: controller.show_frame(p06ValidationView))
        # nextbutton.place(relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )
        submit.bind("<Enter>", mousemove)

        exitbutton = tk.Button(submit, text="Start Over", font=('Helvetica', 16), bg='#004a93', fg="white",
                               command=lambda: controller.show_frame(p07commentpage) )
        exitbutton.place(relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        exitbutton.bind("<Button-1>", unlock)

        instructionsbutton = tk.Button(submit, text="View\n Instructions", font=('Helvetica', 16), bg='#004a93',
                                       fg="white", command= launchinsts)
        instructionsbutton.place(relx=0.0175, rely=.444, width=175, height=75, )

        global hiddenbutton
        hiddenbutton = tk.Button(submit, command=   lambda:  controller.show_frame(p09SaveComments) )
        hiddenbutton.place(relx=0, rely=0, width=0, height=0, )


class p09SaveComments(tk.Frame):
    """
    Save comments page.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        global surveydata, survbox1, survbox2, survbox3

        surveydata = {}

        def nextpopup(self):
            nextbutton.place(relx=0.395, rely=.6575, relwidth=0.18, relheight=0.07, )

        def capturedata():
            surveydata ["Business Unit"] = (survbox1.get(),"comment na" , datetime.now().strftime("%m%d%Y%H%M%S"))
            surveydata ["User Name"] = (survbox2.get(),"comment na" , datetime.now().strftime("%m%d%Y%H%M%S"))
            surveydata ["User Email"] = (survbox3.get(),"comment na" , datetime.now().strftime("%m%d%Y%H%M%S") )

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Collection Survey", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        label2 = tk.Label(submit, text="Please tell us which BU you represent for and how to contact you.", font=('Helvetica', 18), bg='#004a93', fg="white", )
        label2.place(relx=0.22, rely=0.24, relwidth=0.55, relheight=0.07)

        survbox1 = Entry(submit, width=300)
        survbox1.place(relx=0.265, rely=0.35, relwidth=0.45, relheight=0.07)
        survbox1.insert(0, "Please replace this text with your Business Unit name.")
        survbox1.bind("<Button-1>", nextpopup)

        survbox2 = Entry(submit, width=300)
        survbox2.place(relx=0.265, rely=0.45, relwidth=0.45, relheight=0.07)
        survbox2.insert(0, "Please replace this text with your preferred name.")
        survbox2.bind("<Button-1>", nextpopup)

        survbox3 = Entry(submit, width=300)
        survbox3.place(relx=0.265, rely=0.55, relwidth=0.45, relheight=0.07)
        survbox3.insert(0, "Please replace this text with your Allianz email address.")
        survbox3.bind("<Button-1>", nextpopup)

        nextbutton = tk.Button(submit, width=0, height=0, text="Save responses\nand start Survey", font=('Helvetica', 16), bg='#ff7300', fg="black",
                               command=combine_funcs(capturedata, lambda: controller.show_frame(p10SurveyOne)))


class p10SurveyOne(tk.Frame):
    """
    Survey Page 1.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        def nextpopup(self, name):
            global resp
            nextbutton.place(relx=0.395, rely=.765, relwidth=0.18, relheight=0.07, )
            respbox1.place(relx=0.315, rely=0.675, relwidth=0.35, relheight=0.07)
            recresp(self, name)
            respbox1.config(text=resp)

        def capturedata():
            surveydata ["Mission"] = (respbox1["text"], survbox1.get() , datetime.now().strftime("%m%d%Y%H%M%S"))

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Collection Survey", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        label = tk.Label(submit, text="I understand why the Global Portfolio Monitoring team\n"
                                      " is asking for this data and what they are trying to accomplish.",
                         font=('Helvetica', 18), bg='#ff7300', fg="black", )
        label.place(relx=0.22, rely=0.24, relwidth=0.55, relheight=0.07)


        def recresp(self, name):
            global resp, rbutton
            resp = f'Current response: {name["text"]}'

        global survbuttona1
        survbuttona1 = tk.Button(submit, text="The goal of the project is clear.",
                               font=('Helvetica', 14), bg='#008450', fg="white",  command=lambda: nextpopup(self, survbuttona1))
        survbuttona1.place(relx=0.265, rely=0.33, relwidth=0.45, relheight=0.07)


        global survbuttona2
        survbuttona2 = tk.Button(submit, text="This request feels routine and/or redundant. \n(we "
                                             "want to know if this is duplicated work)",
                                font=('Helvetica', 14), bg='#EFB700', fg="black" , command=lambda: nextpopup(self, survbuttona2))
        survbuttona2.place(relx=0.265, rely=0.4125, relwidth=0.45, relheight=0.07)


        global survbuttona3
        survbuttona3 = tk.Button(submit, text="This request feels arbitrary.\n"
                                             " (Is there a reason not to ask you for this?)",
                    font=('Helvetica', 14), bg='#B81D13', fg="white" ,  command=lambda: nextpopup(self, survbuttona3))
        survbuttona3.place(relx=0.265, rely=0.5, relwidth=0.45, relheight=0.07)


        survbox1 = Entry(submit, width=100)
        survbox1.place(relx=0.265, rely=0.59, relwidth=0.45, relheight=0.07)
        survbox1.insert(0, "(Optional) Please replace this text with anything we should know about the request overall.")

        respbox1 = Label(submit, width=100)

        nextbutton = tk.Button(submit, text="Save Response\nand Move On", font=('Helvetica', 16), bg='#ff7300', fg="black",
                               command=combine_funcs(capturedata, lambda: controller.show_frame(p11SurveyTwo)))


class p11SurveyTwo(tk.Frame):
    """
    Survey Page 2.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        def nextpopup(self, name):
            global resp
            nextbutton.place(relx=0.395, rely=.765, relwidth=0.18, relheight=0.07, )
            respbox1.place(relx=0.315, rely=0.675, relwidth=0.35, relheight=0.07)
            recresp(self, name)
            respbox1.config(text=resp)

        def capturedata():
            surveydata["Relevance"] = (respbox1["text"],  survbox1.get(), datetime.now().strftime("%m%d%Y%H%M%S"))

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Collection Survey", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        label = tk.Label(submit, text="The template as provided is a relevant view of the business.",
                         font=('Helvetica', 18), bg='#ff7300', fg="black", )
        label.place(relx=0.22, rely=0.24, relwidth=0.55, relheight=0.07)


        def recresp(self, name):
            global resp, rbutton
            resp = f'Current response: {name["text"]}'

        global survbuttonb1
        survbuttonb1 = tk.Button(submit, text="I understand this request and it aligns with how we manage the business."
                        "\n (Effort and system limitations are addressed in coming questions).", font=('Helvetica', 14),
                        bg='#008450', fg="white",  command=lambda: nextpopup(self, survbuttonb1))
        survbuttonb1.place(relx=0.265, rely=0.33, relwidth=0.45, relheight=0.07)


        global survbuttonb2
        survbuttonb2 = tk.Button(submit, text="I understand the request, but it is missing one or more critical "
                            "elements.\n(Please explain in the comment box.)", font=('Helvetica', 14), bg='#EFB700',
                                 fg="black" , command=lambda: nextpopup(self, survbuttonb2))
        survbuttonb2.place(relx=0.265, rely=0.4125, relwidth=0.45, relheight=0.07)


        global survbuttonb3
        survbuttonb3 = tk.Button(submit, text="This request is not a relevant view of the business\nas we manage it."
                                             " (Please explain in the comment box.)",
                    font=('Helvetica', 14), bg='#B81D13', fg="white" ,  command=lambda: nextpopup(self, survbuttonb3))
        survbuttonb3.place(relx=0.265, rely=0.5, relwidth=0.45, relheight=0.07)


        survbox1 = Entry(submit, width=100)
        survbox1.place(relx=0.265, rely=0.59, relwidth=0.45, relheight=0.07)
        survbox1.insert(0, "(Optional) Please replace this text with anything we should understand about local "
                           "profitability or related metrics.")

        respbox1 = Label(submit, width=100)

        nextbutton = tk.Button(submit, text="Save Response\nand Move On", font=('Helvetica', 16), bg='#ff7300',
                              fg="black", command=combine_funcs(capturedata, lambda: controller.show_frame(p12SurveyThree)))


class p12SurveyThree(tk.Frame):
    """
    Survey Page 3.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        def nextpopup(self, name):
            global resp
            nextbutton.place(relx=0.395, rely=.765, relwidth=0.18, relheight=0.07, )
            respbox1.place(relx=0.315, rely=0.675, relwidth=0.35, relheight=0.07)
            recresp(self, name)
            respbox1.config(text=resp)

        def capturedata():
            surveydata["User Data Access"] = (respbox1["text"],  survbox1.get(), datetime.now().strftime("%m%d%Y%H%M%S"))

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Collection Survey", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        label = tk.Label(submit, text="The data that was requested is available to me\n(you"
                                             " or someone who worked on the template with you).",
                         font=('Helvetica', 18), bg='#ff7300', fg="black", )
        label.place(relx=0.22, rely=0.24, relwidth=0.55, relheight=0.07)


        def recresp(self, name):
            global resp, rbutton
            resp = f'Current response: {name["text"]}'

        global survbuttonc1
        survbuttonc1 = tk.Button(submit, text="I have the data systems I need to complete the request.",
                               font=('Helvetica', 14), bg='#008450', fg="white",  command=lambda: nextpopup(self, survbuttonc1))
        survbuttonc1.place(relx=0.265, rely=0.33, relwidth=0.45, relheight=0.07)

        global survbuttonc2
        survbuttonc2 = tk.Button(submit, text="I have what I need but it is very difficult to complete.",
                                font=('Helvetica', 14), bg='#EFB700', fg="black" , command=lambda: nextpopup(self, survbuttonc2))
        survbuttonc2.place(relx=0.265, rely=0.4125, relwidth=0.45, relheight=0.07)

        global survbuttonc3
        survbuttonc3 = tk.Button(submit, text="I cannot retrieve the data as requested.",
                    font=('Helvetica', 14), bg='#B81D13', fg="white" ,  command=lambda: nextpopup(self, survbuttonc3))
        survbuttonc3.place(relx=0.265, rely=0.5, relwidth=0.45, relheight=0.07)

        survbox1 = Entry(submit, width=100)
        survbox1.place(relx=0.265, rely=0.59, relwidth=0.45, relheight=0.07)
        survbox1.insert(0, "(Optional) Please replace this text with anything we should know about data availability.")

        respbox1 = Label(submit, width=100)

        nextbutton = tk.Button(submit, text="Save Response\nand Move On", font=('Helvetica', 16), bg='#ff7300', fg="black",
                               command=combine_funcs(capturedata, lambda: controller.show_frame(p13SurveyFour)))


class p13SurveyFour(tk.Frame):
    """
    Survey Page 4.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        def nextpopup(self, name):
            global resp
            nextbutton.place(relx=0.395, rely=.765, relwidth=0.18, relheight=0.07, )
            respbox1.place(relx=0.315, rely=0.675, relwidth=0.35, relheight=0.07)
            recresp(self, name)
            respbox1.config(text=resp)

        def capturedata():
            surveydata["Data Infrastructure"] = (respbox1["text"],  survbox1.get(), datetime.now().strftime("%m%d%Y%H%M%S"))

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Collection Survey", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        label = tk.Label(submit, text="The data that was requested is available to someone\n"
                                      "(yourself or any other Allianz employee).",
                         font=('Helvetica', 18), bg='#ff7300', fg="black", )
        label.place(relx=0.22, rely=0.24, relwidth=0.55, relheight=0.07)

        def recresp(self, name):
            global resp, rbutton
            resp = f'Current response: {name["text"]}'

        global survbuttond1
        survbuttond1 = tk.Button(submit, text="All data are available to me OR I know of other Allianz employees\n"
                                             "that have access to data not included in this request.",
                               font=('Helvetica', 14), bg='#008450', fg="white",  command=lambda: nextpopup(self, survbuttond1))
        survbuttond1.place(relx=0.265, rely=0.33, relwidth=0.45, relheight=0.07)

        global survbuttond2
        survbuttond2 = tk.Button(submit, text="I am unsure if data I could not provide is available to other employees.",
                                font=('Helvetica', 14), bg='#EFB700', fg="black" , command=lambda: nextpopup(self, survbuttond2))
        survbuttond2.place(relx=0.265, rely=0.4125, relwidth=0.45, relheight=0.07)


        global survbuttond3
        survbuttond3 = tk.Button(submit, text="I am certain other Allianz employees\n"
                                             "have data not included in this request.",
                    font=('Helvetica', 14), bg='#B81D13', fg="white" ,  command=lambda: nextpopup(self, survbuttond3))
        survbuttond3.place(relx=0.265, rely=0.5, relwidth=0.45, relheight=0.07)

        survbox1 = Entry(submit, width=100)
        survbox1.place(relx=0.265, rely=0.59, relwidth=0.45, relheight=0.07)
        survbox1.insert(0, "(Optional) Please replace this text with anything we should know about who this request "
                           "should be completed by locally.")

        respbox1 = Label(submit, width=100)

        nextbutton = tk.Button(submit, text="Save Response\nand Move On", font=('Helvetica', 16), bg='#ff7300',
                               fg="black",
                               command=combine_funcs(capturedata, lambda: controller.show_frame(p14SurveyFive)))


class p14SurveyFive(tk.Frame):
    """
    Survey Page 5.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        def nextpopup(self):
            nextbutton.place(relx=0.395, rely=.6575, relwidth=0.18, relheight=0.07, )


        def capturedata():
            surveydata["Level of Effort"] = (survbox1.get(), "comment na" , datetime.now().strftime("%m%d%Y%H%M%S"))
            surveydata["Top Issues"] = (survbox2.get(), "comment na" , datetime.now().strftime("%m%d%Y%H%M%S"))
            surveydata["All Other Issues"] = (survbox3.get(), "comment na" , datetime.now().strftime("%m%d%Y%H%M%S"))

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        label = tk.Label(submit, text="Data Collection Survey", font=('Helvetica', 22), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        label = tk.Label(submit, text="Freeform Questions", font=('Helvetica', 18), bg='#004a93', fg="white", )
        label.place(relx=0.26, rely=.1365, relwidth=0.475, relheight=0.08)

        survbox1 = Entry(submit, width=300)
        survbox1.place(relx=0.265, rely=0.35, relwidth=0.45, relheight=0.07)
        survbox1.insert(0, "Please replace this text with an estimate of how many total\n"
                           " hours have been spent on this request to this point.")
        survbox1.bind("<Button-1>", nextpopup)

        survbox2 = Entry(submit, width=300)
        survbox2.place(relx=0.265, rely=0.45, relwidth=0.45, relheight=0.07)
        survbox2.insert(0, "If there is one thing you would like us to know about this request\n "
                           ", please replace this text with that feedback or advice.")
        survbox2.bind("<Button-1>", nextpopup)

        survbox3 = Entry(submit, width=300)
        survbox3.place(relx=0.265, rely=0.55, relwidth=0.45, relheight=0.07)
        survbox3.insert(0,  "If there is anything further you would like us to know about the level of effort\n "
                           "; please replace this text with that information or advice.")
        survbox3.bind("<Button-1>", nextpopup)

        nextbutton = tk.Button(submit, text="Save Response\nand Move On", font=('Helvetica', 16), bg='#ff7300',
                               fg="black",
                               command=combine_funcs(capturedata, lambda: controller.show_frame(p15ExitPage)))


class p15ExitPage(tk.Frame):
    """
    Survey Page 6.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        submit = tk.Frame(self, height=850, width=1150, bg='#dedede', )
        submit.pack(fill="both", expand=True)

        def saveandquit():
            global surveydata

            file = f'us_survey{datetime.now().strftime("%m%d%Y%H%M%S")}.json'
            filepath = os.path.join(str(attachpath), str(file))
            surveydf = pd.DataFrame.from_dict(surveydata.copy(), orient="columns" )
            surveydf.to_json(filepath, orient="table")
            openreportbuttonaction()
            finishbuttonaction()
            sys.exit()

        background_image = tk.PhotoImage(file=bgimage)
        background_image.image = 31
        background_label = tk.Label(submit, image=background_image, height=900, width=1400, )
        background_label.image = background_image
        background_label.place(relwidth=1, relheight=1)

        nextbutton = tk.Button(submit, text="Save and Quit", bg='#ff7300', font=('Helvetica', 22), fg="black",
                         command= saveandquit)
        nextbutton.place(relx=0.38, rely=.5, relwidth=0.2, relheight=0.125, )

        label = tk.Label(submit, text='Thank you for completing the 2020 Third Quarter GPM Data Collection.\n'
                                      'Please review the Excel report. GPM encourages you to reload your data if this submission can be improved.\n'
                                      'If this is the best and final draft, the last step is to send the output to GPM. Email any GPM team\n'
                                      ' member if you have questions about this.\nThank you once again.',
                         font=('Helvetica', 18), bg='#004a93', fg="white", )
        label.place(relx=0.075, rely=0.24, relwidth=0.85, relheight=0.18)


app = root()
app.mainloop()

"""
standard tkinter lines that allow the script to run on loop as an application
"""

if __name__ == '__main__':

    app

"""execution gaurd """