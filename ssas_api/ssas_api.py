# -*- coding: utf-8 -*-
"""
Created on Wed Sep 20 16:59:43 2017

@author: Yehoshua
"""

import os
import sys

import pandas as pd
import clr # name for pythonnet

# logging portion
import logging
logger = logging.getLogger(__name__)

           
def _load_assemblies():
    '''Loads required assemblies, called after function definition.'''

    # Full path of .dll files, gets newest version if there are several folders
    amo_root = (r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
                r"\Microsoft.AnalysisServices.Tabular")
    amo_folder = max(os.listdir(amo_root))
    amo_path = os.path.join(amo_root,amo_folder,
                            os.listdir(os.path.join(amo_root,amo_folder))[0]
                            )
    
    adomd_root = (r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
                  r"\Microsoft.AnalysisServices.AdomdClient")
    adomd_folder = max(os.listdir(adomd_root))
    adomd_path = os.path.join(adomd_root, adomd_folder,
                              os.listdir(os.path.join(adomd_root,adomd_folder))[0]
                              )
    
    # Check that file paths are valid
    assert os.path.isfile(amo_path), ('The filepath for the AMO dll is '
                         'invalid , you passed the filepath: {}'.format(amo_path))
    assert os.path.isfile(adomd_path), ('The filepath for the ADOMD dll is '
                         'invalid , you passed the filepath: {}'.format(adomd_path))

    # load .Net assemblies
    logger.info('Loading .Net assemblies...')
    clr.AddReference('System')
    clr.AddReference('System.Data')
    clr.AddReference(amo_path)
    clr.AddReference(adomd_path)
    
    # Only after loaded .Net assemblies
    global System, DataTable, AMO, ADOMD
    
    import System
    from System.Data import DataTable
    import Microsoft.AnalysisServices.Tabular as AMO
    import Microsoft.AnalysisServices.AdomdClient as ADOMD
    logger.info('Successfully loaded these .Net assemblies: ')

    # log each assembly loaded
    for a in clr.ListAssemblies(True):
        logger.debug(a.split(',')[0])
    

_load_assemblies()


def set_conn_string(ssas_server,db_name,username,password):
    '''
    Sets connection string to SSAS database, in this case designed for Azure Analysis Services
    '''
    conn_string = (
            'Provider=MSOLAP;Data Source={};Initial Catalog={};User ID={};'
            'Password={};Persist Security Info=True;Impersonation Level=Impersonate'
            .format(ssas_server,db_name,username,password)
    )
    return conn_string


def get_DAX(connection_string,dax_string,remove_brackets=False):
    '''
    Executes DAX query and returns the results as a pandas DataFrame
    
    Parameters
    ---------------
    connection_string : string
        Valid SSAS connection string, use the set_conn_string() method to set
    dax_string : string
        Valid DAX query, beginning with EVALUATE or VAR or DEFINE
    remove_brackets : boolean
        If True, then removes brackets from column names

    Returns
    ----------------
    pandas DataFrame with the results
    '''
    dataadapter = ADOMD.AdomdDataAdapter(dax_string,connection_string)
    table = DataTable()
    logger.info('Getting DAX query...')
    try:
        dataadapter.Fill(table)
    except Exception as ex:
        logger.exception('Exception occured:\n')
        sys.exit(1)
    
    col_names = []
    for c in table.Columns.List:
        col_names.append(c.ColumnName)

    num_rows = table.Rows.Count
    
    d = {}
    
    for r in range(num_rows):
        row_dict = {}
        for c in table.Columns.List:
            if isinstance(table.Rows[r][c],System.DBNull):
                row_dict[c.ColumnName] = None
            elif isinstance(table.Rows[r][c],System.DateTime):
                row_dict[c.ColumnName] = pd.Timestamp(table.Rows[r][c].ToShortDateString())
            else:
                row_dict[c.ColumnName] = table.Rows[r][c]
        d[r] = row_dict
    df = pd.DataFrame.from_dict(d,orient='index')

    # remove brackets from column names
    if remove_brackets:
        def remove_brackets(col):
            x = col
            if x[0] == '[':
                x = x.replace('[','')
            x = x.replace('[','_')
            x = x.replace(']','')
            return x
        df = df.rename(columns=remove_brackets)        
    logger.info('DAX query successfully retrieved')
    return df


def process_database(connection_string,db_name,refresh_type='full'):
    '''
    Processes SSAS data model to get new data from underlying database.
    
    Parameters
    -------------
    connection_string : string
        Valid SSAS connection string, use the set_conn_string() method to set
    db_name : string
        The data model on the SSAS server to process
    refresh_type : string, default `full`
        Type of refresh to process. Currently only supports `full`.
    '''
    #connect to the AS instance from Python
    AMOServer = AMO.Server()
    logger.info('Connecting to database...')
    AMOServer.Connect(connection_string)
    
    # Dict of refresh types, I only needed `full`, can add more
    refresh_dict = {'full': AMO.RefreshType.Full}
    
    db = AMOServer.Databases[db_name]
    logger.info('Processing database with refresh type "{}"...'
                .format(refresh_type))
    db.Model.RequestRefresh(refresh_dict[refresh_type])
    op_result = db.Model.SaveChanges()
    if op_result.Impact.IsEmpty:
        logger.info('No objects affected by the refresh')
    
    logger.info('Disconnecting from Database...')
    # Disconnect
    AMOServer.Disconnect()
    