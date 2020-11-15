# -*- coding: latin1 -*-
import arcpy
import numpy

from arcpy.da import NumPyArrayToTable
from pandas import ExcelFile, read_excel


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Util_DIGEOP"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [TableToAlias, ExcelToAlias]


class Tool(object):
    canRunInBackground = False

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Tool"
        self.description = ""

    def getParameterInfo(self):
        """Define parameter definitions"""
        params = None
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
        return


class TableToAlias(Tool):
    category = "Alias"

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Table to Alias"
        self.description = "Lê uma tabela do ArcGIS e preenche os Alias de um Geodatabase"

    def getParameterInfo(self):
        """Define parameter definitions"""
        # Parametro 1 : Feature Class ou Table
        param1 = arcpy.Parameter(
            displayName="Feature Class ou Table para modificação",
            name="in_fc",
            datatype=["DEFeatureClass", "DETable"],
            parameterType="Required",
            direction="Input"
        )

        # Parametro 2 : Table com Campos e Alias
        param2 = arcpy.Parameter(
            displayName="Tabela com Field Names e Alias",
            name="in_table_alias",
            datatype="DETable",
            parameterType="Required",
            direction="Input"
        )

        # Parametro 3 : Coluna com as informações de Field Name
        param3 = arcpy.Parameter(
            displayName="Coluna com os nomes de 'Field Name'",
            name="in_field_name",
            datatype="Field",
            parameterType="Required",
            direction="Input"
        )

        param3.parameterDependencies = [param2.name]
        param3.filter.list = ["Text"]

        # Parametro 4 : Coluna com as informações de Field Alias
        param4 = arcpy.Parameter(
            displayName="Coluna com as informações do 'FieldAlias'",
            name="in_field_alias",
            datatype="Field",
            parameterType="Required",
            direction="Input"
        )
        param4.parameterDependencies = [param2.name]
        param4.filter.list = ["Text"]

        return [param1, param2, param3, param4]

    def execute(self, parameters, messages):
        """The source code of the tool."""
        in_fc = parameters[0].valueAsText
        in_table_alias = parameters[1].valueAsText
        in_field_name = parameters[2].valueAsText
        in_field_alias = parameters[3].valueAsText

        with arcpy.da.SearchCursor(in_table_alias, [in_field_name, in_field_alias]) as cursor:
            for linha in cursor:
                _linha = linha[0].lower().strip()

            campos = [f.name for f in arcpy.ListFields(in_fc)]

            if _linha in campos:
                arcpy.AlterField_management(in_table=in_fc, field=_linha, new_field_alias=linha[1].strip())

        return


class ExcelToAlias(TableToAlias):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Excel To Alias"
        self.description = "Lê uma planilha Excel"

    def getParameterInfo(self):
        # Parametro 2 : Table com Campos e Alias
        param2 = arcpy.Parameter(
            displayName="Arquivo Excel com campos e alias",
            name="in_excel_file",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        param2.filter.list = ["xls", "xlsx"]

        param3 = arcpy.Parameter(
            displayName="Selecione planilha",
            name="in_excel_sheet_name",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        param4 = arcpy.Parameter(
            displayName="Table com os dados extraídos do excel",
            name="in_extracted_table",
            datatype="DETable",
            parameterType="Derived",
            direction="Input"
        )

        # Instanciar os parâmetros da classe pai
        params = super(ExcelToAlias, self).getParameterInfo()

        params[-2].parameterDependencies = [param4.name]
        params[-1].parameterDependencies = [param4.name]

        return [params[0], param2, param3, param4] + params[2:]

    def updateParameters(self, parameters):
        param_in_excel_file = parameters[1]
        param_in_excel_sheet_name = parameters[2]
        param_out_extracted_table = parameters[3]

        if param_in_excel_file.valueAsText:
            workbook = ExcelFile(param_in_excel_file.valueAsText)
            sheets = workbook.sheet_names

            # preenche o objeto Parameter para planilhas, colocando as planilhas disponíveis
            param_in_excel_sheet_name.filter.list = sheets
            param_in_excel_sheet_name.value = sheets[0]

            # Cria table temporária para recuperar os dados do excel
            tmp_table = arcpy.CreateUniqueName("tmp_excel", workspace="in_memory")
            tmp_df = read_excel(workbook, sheetname=param_in_excel_sheet_name.valueAsText)

            # Exporta para Numpy e transforma em ArcGIS Table para extrair as colunas
            NumPyArrayToTable(
                numpy.rec.fromrecords(tmp_df.values, names=tmp_df.columns.tolist()),
                tmp_table
            )
            param_out_extracted_table.value = tmp_table

        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        # Passa os parâmetros necessários para a execução do parent. Indexes [0, 3, 4, 5]
        super(ExcelToAlias, self).execute([parameters[0]] + parameters[3:], messages)
