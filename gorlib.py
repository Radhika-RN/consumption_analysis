import pandas as pd
import random
import numpy as np
import numexpr as ne

def load_file(fpath):
    """
    Loads the excel file and returns a dict of
    sheets
    """
    xls = pd.ExcelFile(fpath)
    sheetdict={}
    for sheet in xls.sheet_names:
        sheetdict[sheet] = pd.read_excel(xls, sheet)
    return sheetdict

def calc_meter_cost(forecast, rates, mid, ezone, aq, sdate, edate):
    """
    Calculates the total cost and consumption for a single meter
    """
    fdsel = (forecast["Meter ID"]== mid) & (forecast["Date"] >= sdate) & (forecast["Date"] < edate)
    rsel = (rates['Exit Zone'] == ezone) & (rates['Annual Quantity (Min)'] <= aq) & (rates['Annual Quantity (Max)'] > aq)
    dff = forecast.loc[fdsel]
    dff = dff.set_index('Date')
    dff = dff.sort_index()
    dfr = rates.loc[rsel, ['Date','Rate (p/kWh)']]
    dfr = dfr.set_index('Date')
    dfr = dfr.sort_index()
    dfr = dfr.reindex(labels=pd.date_range(dfr.index[0],dfr.index[-1],freq='D'), method='pad')
    dfj = dff.join(dfr)
    dfj["Cost"] = dfj['Rate (p/kWh)'] * dfj['kWh']
    return {"Tcost": (dfj["Cost"].sum()/100.0).round(2), "TkWh": dfj["kWh"].sum().round(2)} 


def calc_meter_cost_opt(forecast, rates, mid, ezone, aq, sdate, edate):
    """
    Same as calc_meter_cost: Used for testing optimization
    """
    fdsel = (forecast["Meter ID"]== mid) & (forecast["Date"] >= sdate) & (forecast["Date"] < edate)
    rsel = (rates['Exit Zone'] == ezone) & (rates['Annual Quantity (Min)'] <= aq) & (rates['Annual Quantity (Max)'] > aq)
    dff = forecast.loc[fdsel]
    dff = dff.set_index('Date')
    dff = dff.sort_index()
    dfr = rates.loc[rsel, ['Date','Rate (p/kWh)']]
    dfr = dfr.set_index('Date')
    dfr = dfr.sort_index()
    dfr = dfr.reindex(labels=pd.date_range(dfr.index[0],dfr.index[-1],freq='D'), method='pad')
    #dfj = dff.join(dfr)
    dfr["Cost"] = dfr['Rate (p/kWh)'] * dff['kWh']
    return [(dfr["Cost"].sum()/100.0).round(2), dff["kWh"].sum().round(2)]

def calc_costs(dfconsm, dfrates, dfmeters, sdate, edate):
    """
    Calculates the total cost and consumption for all meters
    """
    mlist=[]
    Tcost=[]
    TkWh=[]
    for index, row in dfmeters.iterrows():
        res = calc_meter_cost(dfconsm, dfrates, row["Meter ID"], row["Exit Zone"],
                         row["Annual Quantity (kWh)"], sdate, edate)
        mlist.append(row["Meter ID"])
        Tcost.append(res["Tcost"])
        TkWh.append(res["TkWh"])
    data = list(zip(mlist,TkWh,Tcost))
    df = pd.DataFrame(data, columns=["Meter ID","Total Estimated Consumption (kWh)","Total Cost (Pounds)"])
    return df 

def calc_costs_opt(dfconsm, dfrates, dfmeters, sdate, edate):
    """
    Same as calc_costs: Used for testing optimization
    """
    dfmeters[["Total Estimated Consumption (kWh)","Total Cost (Pounds)"]] = dfmeters.apply(lambda row: calc_meter_cost_opt(dfconsm, 
                                                                           dfrates, row["Meter ID"], row["Exit Zone"],
                                                                           row["Annual Quantity (kWh)"], sdate, edate), axis=1,result_type="expand")
    return dfmeters[["Meter ID","Total Estimated Consumption (kWh)","Total Cost (Pounds)"]]

def gen_rand_meters(rates, mcnt=100):
    """
    Generates random meter dataframe
    """
    uexits = rates["Exit Zone"].unique().tolist()
    df = pd.DataFrame()
    df["Meter ID"] = random.sample(range(1000,1000+mcnt),mcnt) 
    df["Exit Zone"] = random.choices(uexits, k=mcnt) 
    df["Annual Quantity (kWh)"] = np.random.randint(0,73200*3,size=mcnt) 
    return df

def gen_mock_consn(mlist, sdate, edate):
    """
    Generates mock consumption dataframe
    """
    dflist=[]
    for meter in mlist:
        d_df = pd.DataFrame()
        d_df["Date"] = pd.date_range(sdate,edate,freq='D')
        d_df["Meter ID"] = meter
        d_df["kWh"] = np.random.uniform(2,300,size=d_df.shape[0])
        dflist.append(d_df)
    return pd.concat(dflist)

