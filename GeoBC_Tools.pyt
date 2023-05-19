# -*- coding: utf-8 -*-

import arcpy
import logging
import os
import sys
import ctypes
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from ctypes import wintypes
from datetime import datetime as dt


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "GeoBC_Tools"
        self.alias = "GeoBC Tools"

        # List of tool classes associated with this toolbox
        self.tools = [VolumeCalculator]


class VolumeCalculator(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Volume Calculator"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        aoi = arcpy.Parameter(name='aoi', 
                              displayName='Area of interest polygon layer',
                              direction='Input', 
                              datatype='GPFeatureLayer', 
                              parameterType='Required'
                              )
        fld = arcpy.Parameter(name='fld',
                              displayName='Unique identifier field(s)',
                              direction='Input',
                              datatype='Field',
                              parameterType='Required',
                              multiValue=True,
                              )
        spc = arcpy.Parameter(name='spc',
                              displayName='List of species codes to calculate volumes on',
                              direction='Input',
                              datatype='GPString',
                              parameterType='Optional',
                              multiValue=True
                              )
        xls = arcpy.Parameter(name='xls',
                              displayName='Output Excel document',
                              direction='Output',
                              datatype='DEFile',
                              parameterType='Required'
                              )
        b_un = arcpy.Parameter(name='b_un',
                               displayName='BCGW Username',
                               direction='Input',
                               datatype='GPString',
                               parameterType='Required'
                               )
        
        b_pw = arcpy.Parameter(name='b_pw',
                               displayName='BCGW Password',
                               direction='Input',
                               datatype='GPStringHidden',
                               parameterType='Required'
                               )
        
        fc = arcpy.Parameter(name='fc',
                             displayName='Ouput feature class',
                             direction='Output',
                             datatype='DEFeatureClass',
                             parameterType='Optional'
                             )
        vri = arcpy.Parameter(name='vri',
                              displayName='Optional VRI dataset to use instead of current',
                              direction='Input',
                              datatype='GPFeatureLayer',
                              parameterType='Optional'
                              )

        fld.parameterDependencies = [aoi.name]
        xls.filter.list = ['xlsx']
        params = [aoi, fld, spc, xls, b_un, b_pw, fc, vri]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        if arcpy.ProductInfo() == 'ArcView':
            arcpy.AddError('You do not have the necessary licence level to use this tool. Please ensure you have at least ArcEditor activated')
            return
        volume = CalculateVolumes(
            aoi=parameters[0].valueAsText, 
            id_fields=parameters[1].valueAsText,
            bcgw_un=parameters[4].valueAsText,
            bcgw_pw=parameters[5].valueAsText,
            out_fc=parameters[6].valueAsText,
            excel=parameters[3].valueAsText,
            species_list=parameters[2].valueAsText,
            vri=parameters[7].valueAsText)
        
        lst_species_fields = volume.calculate_volumes()
        volume.create_excel(lst_species_fields=lst_species_fields)
        arcpy.Delete_management('in_memory')
        del volume
        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return


class CalculateVolumes:
    def __init__(self, aoi, id_fields, species_list, excel, bcgw_un, bcgw_pw, out_fc, vri):
        self.aoi = aoi
        self.lst_fld_ids = str(id_fields).split(';') if not isinstance(id_fields, list) else id_fields
        self.lst_species = str(species_list).split(';') if not isinstance(species_list, list) else species_list
        self.lst_species.append('Other')
        self.bcgw_un = bcgw_un
        self.bcgw_pw = bcgw_pw
        self.scratch_gdb = 'in_memory'
        self.fc_volume_summary = None if out_fc == '#' else out_fc
        self.scratch_gdb = 'in_memory'
        self.aprx = arcpy.mp.ArcGISProject('CURRENT')
        self.sde_folder = self.aprx.homeFolder
        self.output_xls = Environment.get_full_path(excel)

        if not self.fc_volume_summary:
            self.fc_volume_summary = os.path.join(self.scratch_gdb, 'volume_summary')
        arcpy.AddMessage(self.fc_volume_summary)
    
        # Connect to SDE databases and create output folders
        self.bcgw_db = Environment.create_bcgw_connection(location=self.sde_folder, bcgw_user_name=self.bcgw_un, bcgw_password=self.bcgw_pw, )
    
        self.__vri = os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.VEG_COMP_LYR_R1_POLY') if not vri  else vri
    
        self.fc_aoi = os.path.join(self.scratch_gdb, 'aoi')
        self.fc_vri_copy = os.path.join(self.scratch_gdb, 'vri_copy')
        self.fc_vri_clip = os.path.join(self.scratch_gdb, 'vri_clip')
        self.fc_vri_aoi = os.path.join(self.scratch_gdb, 'vri_aoi')
    
        self.fld_volha_total = 'VOLHA_TOTAL'
        self.fld_vol_total = 'VOLUME_TOTAL'
        self.fld_species = 'SPECIES_ID'
    
        arcpy.env.extent = arcpy.Describe(self.aoi).extent
        arcpy.env.overwriteOutput = True
    
    def __del__(self):
        Environment.delete_bcgw_connection(location=self.sde_folder)
    
    def calculate_volumes(self):
        arcpy.AddMessage('Calculating volumes')
    
        fld_volha_spc = 'VOLHA_SPC'
        fld_vol_spc = 'VOL_SPC'
        fld_area_ha = 'AREA_HECTARES'
        fld_species_cd = 'SPECIES_CD_'
        fld_species_pct = 'SPECIES_PCT_'
        fld_live_vol = 'LIVE_VOL_PER_HA_SPP'
    
        arcpy.CopyFeatures_management(in_features=self.aoi, out_feature_class=self.fc_aoi)
    
        arcpy.AddMessage('Adding in aoi')
        arcpy.PairwiseIntersect_analysis(in_features=[self.fc_aoi, self.__vri],
                                        out_feature_class=self.fc_vri_aoi, join_attributes='NO_FID')
    
    
        lst_cursor_fields = []
        lst_species_fields = []
        for s in self.lst_species:
            lst_species_fields.extend([f'{s}_AREA', f'{s}_VOLUME', f'{s}_VOLHA', f'{s}_VOLHA_TOTAL'])
    
        lst_fields = [self.fld_volha_total, fld_area_ha, self.fld_vol_total] + lst_species_fields
    
        for i in range(1, 7):
            for fld in [fld_volha_spc, fld_vol_spc]:
                arcpy.AddMessage(f'Adding field {fld}{i}...')
                arcpy.AddField_management(in_table=self.fc_vri_aoi, field_name=f'{fld}{i}', field_type='DOUBLE')
    
        for fld in lst_fields:
            arcpy.AddMessage(f'Adding field {fld}...')
            arcpy.AddField_management(in_table=self.fc_vri_aoi, field_name=fld, field_type='DOUBLE')
    
        for i in range(1, 7):
            lst_cursor_fields += [f'{fld_volha_spc}{i}', f'{fld_vol_spc}{i}',
                                  f'{fld_species_cd}{i}', f'{fld_species_pct}{i}',
                                  f'{fld_live_vol}{i}_125']
    
        arcpy.AddMessage('Calculating species and volume values...')
        lst_cursor_fields += lst_fields + ['SHAPE@AREA'] + self.lst_fld_ids
        lst_species_fields
        with arcpy.da.UpdateCursor(self.fc_vri_aoi, lst_cursor_fields) as u_cursor:
            for row in u_cursor:
                volha_total = 0
                vol_total = 0
                ids = ()
                for fld in self.lst_fld_ids:
                    ids += (row[lst_cursor_fields.index(fld)],)
    
                if not all(ids):
                    u_cursor.deleteRow()
                    continue
    
                for fld in lst_species_fields:
                    row[lst_cursor_fields.index(fld)] = 0
    
                for i in range(1, 7):
                    area = row[lst_cursor_fields.index('SHAPE@AREA')]
                    area_ha = area / 10000
                    species_cd = str(row[lst_cursor_fields.index(f'{fld_species_cd}{i}')])
                    live_vol_125 = row[lst_cursor_fields.index(f'{fld_live_vol}{i}_125')]
    
                    volha_spc = 0
    
                    if species_cd == 'None':
                        break
    
                    volha_spc = live_vol_125
                    if not volha_spc:
                        volha_spc = 0
    
                    vol_spc = volha_spc * area_ha
                    row[lst_cursor_fields.index(f'{fld_volha_spc}{i}')] = volha_spc
                    row[lst_cursor_fields.index(f'{fld_vol_spc}{i}')] = vol_spc
    
                    bl_flag = False
                    for species in self.lst_species:
                        if species_cd.startswith(str(species).upper()):
                            row[lst_cursor_fields.index(f'{species}_VOLUME')] += vol_spc
                            row[lst_cursor_fields.index(f'{species}_AREA')] += area_ha
                            bl_flag = True
    
                    if not bl_flag:
                        row[lst_cursor_fields.index('Other_VOLUME')] += vol_spc
                        row[lst_cursor_fields.index('Other_AREA')] += area_ha
    
                    volha_total += volha_spc
                    vol_total += vol_spc
    
                row[lst_cursor_fields.index(self.fld_volha_total)] = volha_total
                row[lst_cursor_fields.index(fld_area_ha)] = area_ha
                row[lst_cursor_fields.index(self.fld_vol_total)] = vol_total
    
                u_cursor.updateRow(row)
    
        stats_fields = []
        for fld in lst_species_fields + [self.fld_vol_total]:
            stats_fields.append([fld, 'SUM'])
    
        arcpy.AddMessage('Summarizing volumes')
        arcpy.Dissolve_management(in_features=self.fc_vri_aoi, out_feature_class=self.fc_volume_summary,
                                  dissolve_field=self.lst_fld_ids,
                                  statistics_fields=stats_fields)
    
        for fld in lst_species_fields + [self.fld_vol_total]:
            arcpy.AlterField_management(in_table=self.fc_volume_summary, field=f'SUM_{fld}',
                                            new_field_name=fld, new_field_alias=fld)
    
        arcpy.AddField_management(in_table=self.fc_volume_summary, field_name=self.fld_volha_total, field_type='DOUBLE')
        lst_cursor_fields = [self.fld_vol_total, self.fld_volha_total, 'SHAPE@AREA'] + \
                            self.lst_fld_ids + lst_species_fields
    
        with arcpy.da.UpdateCursor(self.fc_volume_summary, lst_cursor_fields) as u_cursor:
            for row in u_cursor:
                row[lst_cursor_fields.index(self.fld_volha_total)] = \
                    row[lst_cursor_fields.index(self.fld_vol_total)] / \
                    (row[lst_cursor_fields.index('SHAPE@AREA')] / 10000)
    
                for s in self.lst_species:
                    try:
                        row[lst_cursor_fields.index(f'{s}_VOLHA')] = \
                            row[lst_cursor_fields.index(f'{s}_VOLUME')] / row[lst_cursor_fields.index(f'{s}_AREA')]
                    except:
                        row[lst_cursor_fields.index(f'{s}_VOLHA')] = 0
                    try:
                        row[lst_cursor_fields.index(f'{s}_VOLHA_TOTAL')] = \
                            row[lst_cursor_fields.index(f'{s}_VOLUME')] / \
                                (row[lst_cursor_fields.index('SHAPE@AREA')] / 10000)
                    except:
                        row[lst_cursor_fields.index(f'{s}_VOLHA_TOTAL')] = 0
    
                ids = ()
                for fld in self.lst_fld_ids:
                    ids += (row[lst_cursor_fields.index(fld)],)
                u_cursor.updateRow(row)
    
        # for fc in [self.fc_vri_copy, vri_lyr]:
        #     arcpy.Delete_management(in_data=fc)
    
        return lst_species_fields
    
    def create_excel(self, lst_species_fields):
        arcpy.AddMessage('Creating excel output')
        lst_fields = self.lst_fld_ids + ['SHAPE@AREA', self.fld_vol_total, self.fld_volha_total] + lst_species_fields
        lst_vol_summary = []
        with arcpy.da.SearchCursor(self.fc_volume_summary, lst_fields) as s_cursor:
            for row in s_cursor:
                lst_item = []
                for i in range(0, len(lst_fields)):
                    if lst_fields[i] == 'SHAPE@AREA':
                        lst_item.append(row[i] / 10000)
                    else:
                        lst_item.append(row[i])
    
                lst_vol_summary.append(lst_item)
    
        lst_species_columns = []
        for s in self.lst_species:
            lst_species_columns.extend([f'{s} Area (ha)', f'{s} Volume (m3)', f'{s} Volume/ha (m3)', f'{s} Total Volume/ha (m3)'])

            volume_columns = self.lst_fld_ids + ['Total Area (ha)', 'Total Volume (m3)', 'Total Volume/ha (m3)'] + \
                                lst_species_columns
        df_vol = pd.DataFrame(data=lst_vol_summary, columns=volume_columns)

        # wb = openpyxl.Workbook()

        # ws = wb.create_sheet(title='Volume Summary')
        # rows = dataframe_to_rows(df_vol)

        # for r_idx, row in enumerate(rows, 1):
        #     for c_idx, value in enumerate(row, 1):
        #         ws.cell(row=r_idx, column=c_idx, value=value)

        # arcpy.AddMessage(self.output_xls)
        # wb.save(filename=self.output_xls)


        with pd.ExcelWriter(path=self.output_xls, engine='openpyxl') as x_writer:


            df_vol.to_excel(excel_writer=x_writer, sheet_name='Volume Summary', index=False)
            ws = x_writer.sheets['Volume Summary']
            wb = x_writer.book
            num_format = '#,##0.00'

            for i in range(len(self.lst_fld_ids) + 1, len(volume_columns) + 1):
                for j in range(2, df_vol.shape[0] + 2):
                    ws.cell(column=i, row=j).number_format = num_format

            for i, width in enumerate(get_col_widths(df_vol)):
                ws.column_dimensions[openpyxl.utils.cell.get_column_letter(i + 1)].width = width + 2
                # ws.set_column(i - 1, i - 1, width)
        
        # wb = openpyxl.load_workbook(filename=self.output_xls)
        # ws = wb['Volume Summary']
        # for i in range(len(self.lst_fld_ids), len(volume_columns)):
        #         ws.column_dimensions[openpyxl.utils.cell.get_column_letter(i + 1)].number_format = num_format
        #         # ws.set_column('{0}:{0}'.format(str(chr(97 + i)).upper()), None, num_format)
    
        # wb.save(filename=self.output_xls)

        
def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right

    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) + 4 for col in dataframe.columns]


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


class Environment:
    """
    ------------------------------------------------------------------------------------------------------------
        CLASS: Contains general environment functions and processes that can be used in python scripts
    ------------------------------------------------------------------------------------------------------------
    """

    def __init__(self):
        pass

    # Set up variables for getting UNC paths
    mpr = ctypes.WinDLL('mpr')

    ERROR_SUCCESS = 0x0000
    ERROR_MORE_DATA = 0x00EA

    wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)
    mpr.WNetGetConnectionW.restype = wintypes.DWORD
    mpr.WNetGetConnectionW.argtypes = (wintypes.LPCWSTR,
                                       wintypes.LPWSTR,
                                       wintypes.LPDWORD)


    @staticmethod
    def create_bcgw_connection(location, bcgw_user_name, bcgw_password, db_name='Temp_BCGW.sde', logger=None):
        """
            ------------------------------------------------------------------------------------------------------------
                FUNCTION: Creates a connection object to the bcgw SDE database

                Parameters:
                    location: path to where the database connection object will be saved
                    bcgw_user_name: User name for the BCGW
                    bcgw_password: Password for the BCGW
                    db_name: database name
                    logger: logging object for message output

                Return: None
            ------------------------------------------------------------------------------------------------------------
        """
        if logger:
            logger.info('Connecting to BCGW')

        if not arcpy.Exists(os.path.join(location, db_name)):
            arcpy.CreateDatabaseConnection_management(out_folder_path=location,
                                                      out_name=db_name[:-4],
                                                      database_platform='ORACLE',
                                                      instance='bcgw.bcgov/idwprod1.bcgov',
                                                      username=bcgw_user_name,
                                                      password=bcgw_password,
                                                      save_user_pass='SAVE_USERNAME')
        return os.path.join(location, 'Temp_BCGW.sde')

    @staticmethod
    def delete_bcgw_connection(location, db_name='Temp_BCGW.sde', logger=None):
        """
           ------------------------------------------------------------------------------------------------------------
               FUNCTION: Deletes the bcgw database connection object

               Parameters:
                   location: path to where the database connection object exists
                   db_name: database name
                   logger: logging object for message output

               Return: None
           ------------------------------------------------------------------------------------------------------------
        """
        bcgw_path = os.path.join(location, db_name)
        if logger:
            logger.info('Deleting BCGW connection')
        if location == 'Database Connections':
            os.remove(Environment.sde_connection(db_name))
        else:
            os.remove(bcgw_path)


    @staticmethod
    def get_network_path(local_name):
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: Take in a drive letter (ie. 'W:') and return the full network path for that letter.
                Mapped drives are not recognized when trying to open a file, thereby requiring the use of this function

            Parameters:
                local_name str: letter and colon for the mapped drive

            Return str: unc value for the mapped drive
        ------------------------------------------------------------------------------------------------------------
        """
        length = (wintypes.DWORD * 1)()
        result = Environment.mpr.WNetGetConnectionW(local_name, None, length)
        if result != Environment.ERROR_MORE_DATA:
            raise ctypes.WinError(result)
        remote_name = (wintypes.WCHAR * length[0])()
        result = Environment.mpr.WNetGetConnectionW(local_name, remote_name, length)
        if result != Environment.ERROR_SUCCESS:
            raise ctypes.WinError(result)
        return remote_name.value

    @staticmethod
    def get_full_path(str_file):
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: determine if a mapped drive file path is being used.
                It then calls a function to replace the drive letter with the full UNC path

            Parameters:
                str_file str: file path that needs to be checked for mapped drives

            Return str: correct file path
        ------------------------------------------------------------------------------------------------------------
        """

        if str_file.startswith('\\'):
            return str_file

        # Check to see if the path is a valid file.  If not it is most likely on a mapped drive
        str_file = str_file.replace("'", "")
        if not os.path.isfile(str_file):
            file_path = os.path.join(Environment.get_network_path(str_file[:2]), str_file[2:])
        else:
            file_path = str_file

        return file_path