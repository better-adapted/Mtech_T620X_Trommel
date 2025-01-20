import pandas as pd
import fileinput
import six
import shutil
import numpy
import re
from datetime import datetime, timezone

# TODO - timestamp the comments in the import!
# TODO - add stuck on power code generation

PRINT_COLS = False

def replace_section_in_file(filename_in,filename_out,Lines,marker):
    found_start_marker = False
    found_stop_marker = False

    start_block=""
    stop_block=""

    start_marker = marker + "_CODE_GEN_START"
    stop_marker = marker + "_CODE_GEN_STOP"

    with open(filename_in,'rt',) as fp:
        # Loop to replace text in place
        for line in fp:
            if line.find(start_marker)>=0:
                found_start_marker = True            
            elif line.find(stop_marker)>=0:
                if found_stop_marker == False:
                    found_stop_marker = True                

            if found_start_marker == False:
                start_block+=line

            if found_stop_marker == True:
                stop_block+=line

        if found_start_marker == False:
            return

        if found_stop_marker == False:
            return

    opfile = open(filename_out, "wt+")
    opfile.write(start_block+start_marker+"\r\n")

    opfile.write("// generated (UTC): " + str(datetime.now(timezone.utc)) + "\r\n")

    for replace_line  in Lines:
        opfile.write(replace_line+"\n")

    opfile.write(stop_block)
    opfile.close()

    print(f'replace {filename_in=} {filename_in=} {marker=} Done')

    return 0

T620X_TEST=False
IMPORT_Filename =""
EXPORT_Filename =""

if T620X_TEST :
    IO_filename = '../T620X_IO.xlsx'

    SRC_EXPORT_FILE = "T620X_HMI_CR1076_SP19p4.export"
    SRC_EXPORT_DIR = "../CR1076_HMI/"

    IMPORT_Directory = SRC_EXPORT_DIR
    IMPORT_Filename = IMPORT_Directory + SRC_EXPORT_FILE

    EXPORT_Directory = IMPORT_Directory + "codegen/"
    EXPORT_Filename = EXPORT_Directory + SRC_EXPORT_FILE.replace("CODE_ORG","CODE_TEST")
    
    Machine_IO_Individual_PLC_Structs=False

shutil.copy(IO_filename, EXPORT_Directory) # keep a copy of the excel file we used

df = pd.read_excel(IO_filename,sheet_name='IO_LIST')

Machine_Link_Inputs_Main_List = []
Machine_Link_Outputs_Main_List = []
Machine_Link_Inputs_Init_List = []
Machine_Link_Outputs_Init_List = []
Machine_IO_Struct_Inputs_Init_List = []
Machine_IO_Struct_Outputs_Init_List = []

Alarms_Monitor_IO_Declare_List = []
Alarms_Monitor_IO_STL_List = []

HMI_IO_Init_Input_Links_LIST = []
HMI_IO_Init_Output_Links_LIST = []
HMI_IO_Update_Inputs_LIST = []
HMI_IO_Update_Outputs_LIST = []

if PRINT_COLS:
    for column_name in df:
        print("column_name: %s,%s"%((column_name),type(column_name)))

Machine_IO_Prefix = "Machine_IO."

Text_Lists_Common_Header = ("TextList" + "\t" + "Id" + "\t" + "Default" + "\t" + "ge" + "\t" + "en")

TextList_IO_Diag = [Text_Lists_Common_Header]
AlarmGroup_App = [Text_Lists_Common_Header]
AlarmGroup_IO = [Text_Lists_Common_Header]

for row in df.index:
    row_PLC_Name = df['PLC_Name'][row]
    row_PLC_REF = df['PLC_REF'][row]
    row_PLC_IO_TAG = df['PLC_IO_TAG'][row]
    row_TYPE = df['TYPE'][row]
    row_Feature_Name = df['Feature_Name'][row]
    row_Connector = df['Connector'][row]
    row_IO_ALARM_ID = df['IO_Alarm_ID'][row]
    row_IO_ALARM_TEXT = df['IO_Alarm_Text'][row]

    # if spare keep consistant for IpCom OpCom
    if re.search(row_Feature_Name,"SPARE",re.IGNORECASE):
        row_Feature_Name = "SPARE_" + row_PLC_REF

    print(row_PLC_Name,row_PLC_REF,row_Feature_Name)

    if isinstance(row_PLC_Name, six.string_types) & isinstance(row_PLC_REF, six.string_types) & isinstance(row_Feature_Name, six.string_types) :

        row_Feature_Name_underscores = row_Feature_Name.replace(".","_")        
        
        if Machine_IO_Individual_PLC_Structs == True:
            if row_TYPE == "IN":
                Machine_IO_Prefix_in_out_temp = row_PLC_Name + "_Inputs."

            if row_TYPE == "OUT":
                Machine_IO_Prefix_in_out_temp = row_PLC_Name + "_Outputs."
        else:
            if row_TYPE == "IN":
                Machine_IO_Prefix_in_out_temp = "IN."

            if row_TYPE == "OUT":
                Machine_IO_Prefix_in_out_temp = "OUT."

        # Machine_IO.DICP_Inputs.CAS_On_Off_Switch
        Machine_IO_Object = Machine_IO_Prefix + Machine_IO_Prefix_in_out_temp + row_Feature_Name_underscores


        # Machine_IO.DICP_Inputs.CAS_On_Off_Switch(); // CR0709.IN0100
        Machine_IO_Main_IO_Ref = Machine_IO_Object + "(); //" + row_PLC_REF

        #print(row_PLC_Name,row_PLC_REF,row_Feature_Name,"=>",Machine_IO_Inputs_Text)
        if row_TYPE == "IN":
            Machine_Link_Inputs_Main_List.append(Machine_IO_Main_IO_Ref)
        if (row_TYPE == "OUT") or (row_TYPE == "OUT_GRP") or (row_TYPE == "OUTA"):
            Machine_Link_Outputs_Main_List.append(Machine_IO_Main_IO_Ref)
            
        # NVL_IO_DICP.All_Inputs.IN0100
        if row_TYPE == "IN":
            Machine_IO_nvl_ref_temp = "NVL_IO_" + row_PLC_Name + ".All_Inputs." + row_PLC_IO_TAG
        
        if row_TYPE == (row_TYPE == "OUT") or (row_TYPE == "OUT_GRP") or (row_TYPE == "OUTA"):
            Machine_IO_nvl_ref_temp = "NVL_IO_" + row_PLC_Name + ".All_Outputs." + row_PLC_IO_TAG

        #  Machine_IO.Node_DICP
        Machine_IO_node_ref_temp = Machine_IO_Prefix + "Node_" + row_PLC_Name

        if row_TYPE == "IN":
            Machine_IO_input_ref_temp = "NVL_IO_" + row_PLC_Name + ".All_Inputs." + row_PLC_IO_TAG
            # Machine_IO.DICP_Inputs.CAS_On_Off_Switch.Init('DICP','CAS_On_Off_Switch','IN0100','a63',NVL_IO_DICP.All_Inputs.IN0100,Machine_IO.Node_DICP);
            Machine_IO_Init_Text_Temp = Machine_IO_Object
            Machine_IO_Init_Text_Temp += ".Init("
            Machine_IO_Init_Text_Temp += "'" + row_PLC_Name + "',"
            Machine_IO_Init_Text_Temp += "'" + row_Feature_Name + "',"
            Machine_IO_Init_Text_Temp += "'" + row_PLC_IO_TAG + "',"
            Machine_IO_Init_Text_Temp += "'" + row_Connector + "',"
            Machine_IO_Init_Text_Temp += Machine_IO_input_ref_temp + ","
            Machine_IO_Init_Text_Temp += Machine_IO_node_ref_temp
            Machine_IO_Init_Text_Temp += ");"
            Machine_Link_Inputs_Init_List.append(Machine_IO_Init_Text_Temp)


        if (row_TYPE == "OUT") or (row_TYPE == "OUT_GRP") or (row_TYPE == "OUTA"):
            # NVL_Outputs_States_DICP.All.OUT0000
            Machine_IO_output_ref_temp = "NVL_Outputs_States_" + row_PLC_Name + ".All." + row_PLC_IO_TAG

            # NVL_IO_DICP.All_Outputs_Diag.OUT0000
            Machine_IO_nvl_op_diag_ref_temp = Machine_IO_nvl_ref_temp = "NVL_IO_" + row_PLC_Name + ".All_Outputs_Diag." + row_PLC_IO_TAG

            # Machine_IO.DICP_Outputs.Engine_Crank.Init('DICP','Engine_Crank','OUT0000','a16',NVL_Outputs_States_DICP.All.OUT0000,NVL_IO_DICP.All_Outputs_Diag.OUT0000,Machine_IO.Node_DICP);
            Machine_IO_Init_Text_Temp = Machine_IO_Object
            Machine_IO_Init_Text_Temp += ".Init("
            Machine_IO_Init_Text_Temp += "'" + row_PLC_Name + "',"
            Machine_IO_Init_Text_Temp += "'" + row_Feature_Name + "',"
            Machine_IO_Init_Text_Temp += "'" + row_PLC_IO_TAG + "',"
            Machine_IO_Init_Text_Temp += "'" + row_Connector + "',"
            Machine_IO_Init_Text_Temp += Machine_IO_output_ref_temp + ","
            Machine_IO_Init_Text_Temp += Machine_IO_nvl_op_diag_ref_temp + ","
            Machine_IO_Init_Text_Temp += Machine_IO_node_ref_temp
            Machine_IO_Init_Text_Temp += ");"
            Machine_Link_Outputs_Init_List.append(Machine_IO_Init_Text_Temp)


        # building up Structs
        if row_TYPE == "IN":
            Machine_IO_Struct_Inputs_Text_Temp = row_Feature_Name_underscores
            Machine_IO_Struct_Inputs_Text_Temp += ": IpCom;"
            Machine_IO_Struct_Inputs_Text_Temp += " // " + row_PLC_Name + "." + row_PLC_IO_TAG
            Machine_IO_Struct_Inputs_Init_List.append(Machine_IO_Struct_Inputs_Text_Temp)

        if (row_TYPE == "OUT") or (row_TYPE == "OUTA"):
            Machine_IO_Struct_Outputs_Text_Temp = row_Feature_Name_underscores
            Machine_IO_Struct_Outputs_Text_Temp += ": OpCom;"
            Machine_IO_Struct_Outputs_Text_Temp += " // " + row_PLC_Name + "." + row_PLC_IO_TAG
            Machine_IO_Struct_Outputs_Init_List.append(Machine_IO_Struct_Outputs_Text_Temp)

        if row_TYPE == "OUT_GRP":
            Machine_IO_Struct_Outputs_Text_Temp = row_Feature_Name_underscores
            Machine_IO_Struct_Outputs_Text_Temp += ": OpGrpCom;"
            Machine_IO_Struct_Outputs_Text_Temp += " // " + row_PLC_Name + "." + row_PLC_IO_TAG
            Machine_IO_Struct_Outputs_Init_List.append(Machine_IO_Struct_Outputs_Text_Temp)

        if (row_TYPE == "IN") or  (row_TYPE == "OUT") or (row_TYPE == "OUT_GRP") or (row_TYPE == "OUTA") :
            # ESTOPS_SAFETY_RELAY: MGSE_AlarmBit;
            Alarms_Monitor_IO_Declare_List.append("\t" + row_Feature_Name_underscores + ": MGSE_AlarmBit;")

            # CAS_On_Off_Switch(LIST:=LIST,Bit_Input:=Machine_IO.DICP_Inputs.CAS_On_Off_Switch.Alarm_Active,ID:=1000,Add_Info:=Machine_IO.DICP_Inputs.CAS_On_Off_Switch.AlmLog.Buffer,Group:=IO_Error_Group);
            Machine_IO_Alm_Temp = row_Feature_Name_underscores
            Machine_IO_Alm_Temp += "(LIST:=LIST,"
            Machine_IO_Alm_Temp += "Bit_Input:=" + Machine_IO_Object + ".Alarm_Active,"
            Machine_IO_Alm_Temp += "ID:=" + str(int(row_IO_ALARM_ID)) + ","
            Machine_IO_Alm_Temp += "Add_Info:=" + Machine_IO_Object + ".AlmLog.Buffer,"
            Machine_IO_Alm_Temp += "Group:=IO_Error_Group);"            
            Alarms_Monitor_IO_STL_List.append(Machine_IO_Alm_Temp)

            # Device.Application.AlarmGroup_IO	1001	ESTOPS_SAFETY_RELAY {CHASSIS.IN0100}	ESTOPS_SAFETY_RELAY {CHASSIS.IN0100}	ESTOPS_SAFETY_RELAY {CHASSIS.IN0100}
            # TABS between! - depends on the number of languages in the system!
            Temp_AlarmGroup_IO_Desc = row_Feature_Name_underscores + " (" + row_PLC_REF + ")"
            Temp_AlarmGroup_IO_Full = "Device.Application.AlarmGroup_IO" + "\t" + str(int(row_IO_ALARM_ID)) + "\t" + Temp_AlarmGroup_IO_Desc + "\t" + Temp_AlarmGroup_IO_Desc + "\t" + Temp_AlarmGroup_IO_Desc
            AlarmGroup_IO.append(Temp_AlarmGroup_IO_Full)

            # Device.Application.TextList_IO_Diag	CHASSIS.IN0100	ESTOPS_SAFETY_RELAY	ESTOPS_SAFETY_RELAY	ESTOPS_SAFETY_RELAY
            # TABS between! - depends on the number of languages in the system!
            Temp_TextList_IO_Diag_Desc = row_Feature_Name_underscores
            Temp_TextList_IO_Diag_Full = "Device.Application.TextList_IO_Diag" + "\t" + row_PLC_REF + "\t" + Temp_TextList_IO_Diag_Desc + "\t" + Temp_TextList_IO_Diag_Desc + "\t" + Temp_TextList_IO_Diag_Desc
            TextList_IO_Diag.append(Temp_TextList_IO_Diag_Full)

        if row_TYPE == "IN":
            # CHASSIS.IN0100.Init(Machine_IO.CHASSIS_Inputs.EStop_Safety_Relay_Signal);
            Temp_HMI_IO_Init_Links = row_PLC_REF + ".Init(" + Machine_IO_Object +");"
            HMI_IO_Init_Input_Links_LIST.append(Temp_HMI_IO_Init_Links)

            # DICP.IN0100();
            Temp_HMI_IO_Init_Links = row_PLC_REF + "();"
            HMI_IO_Update_Inputs_LIST.append(Temp_HMI_IO_Init_Links)


        if (row_TYPE == "OUT") or (row_TYPE == "OUT_GRP") or (row_TYPE == "OUTA"):
            # CHASSIS.OUT0000.Init(Machine_IO.CHASSIS_Outputs.Chassis_AC_Active_LED);
            Temp_HMI_IO_Init_Links = row_PLC_REF + ".Init(" + Machine_IO_Object +");"
            HMI_IO_Init_Output_Links_LIST.append(Temp_HMI_IO_Init_Links)

            # DICP.OUT0000();
            Temp_HMI_IO_Init_Links = row_PLC_REF + "();"
            HMI_IO_Update_Outputs_LIST.append(Temp_HMI_IO_Init_Links)

            print(f'Machine_IO {row_PLC_Name=} {row_PLC_REF=} {row_PLC_IO_TAG=} Done')         

TempFile1 = "temp1.export"
TempFile2 = "temp2.export"
TempFile3 = "temp3.export"
TempFile4 = "temp4.export"
TempFile5 = "temp5.export"
TempFile6 = "temp6.export"
TempFile7 = "temp7.export"
TempFile8 = "temp8.export"
TempFile9 = "temp9.export"
TempFile10 = "temp10.export"
TempFile11 = "temp11.export"
TempFile12 = "temp12.export"
TempFile13 = "temp13.export"
TempFile14 = "temp14.export"
TempFile15 = "temp15.export"

replace_section_in_file(IMPORT_Filename,TempFile1,Machine_Link_Inputs_Main_List,"// Machine_Link_Inputs_Main")
replace_section_in_file(TempFile1,TempFile2,Machine_Link_Inputs_Init_List,"// Machine_Link_Inputs_Init")
replace_section_in_file(TempFile2,TempFile3,Machine_Link_Outputs_Main_List,"// Machine_Link_Outputs_Main")
replace_section_in_file(TempFile3,TempFile4,Machine_Link_Outputs_Init_List,"// Machine_Link_Outputs_Init")

replace_section_in_file(TempFile4,TempFile5,Machine_IO_Struct_Inputs_Init_List,"// Machine_IO_Inputs_STRUCT")
replace_section_in_file(TempFile5,TempFile6,Machine_IO_Struct_Outputs_Init_List,"// Machine_IO_Outputs_STRUCT")

df_alm = pd.read_excel(IO_filename,sheet_name='Alarms')

GVL_Alarms_Declare_List = []

Alarms_Monitor_App_Declare_List = []
Alarms_Monitor_App_STL_List = []

if PRINT_COLS:
    for column_name in df_alm:
        print("column_name: %s,%s"%((column_name),type(column_name)))

for row_alm in df_alm.index:
    row_alm_ID = df_alm['ID#'][row_alm]
    row_alm_TAG = df_alm['TAG'][row_alm]
    row_alm_HUSH = df_alm['HUSH'][row_alm]
    row_alm_Default_Text = df_alm['Default_Text'][row_alm]

    # BOOL bit for each App Alarm - in GVL
    row_alm_temp = "\t" + row_alm_TAG + ": BOOL; // ID=" + str(row_alm_ID)
    GVL_Alarms_Declare_List.append(row_alm_temp)

    # Alarmbit FB for each App Alarm 
    Alarms_Monitor_App_Declare_List.append("\t" + row_alm_TAG + ": MGSE_AlarmBit;")

    if numpy.isnan(row_alm_HUSH):
        row_alm_HUSH=0

    # STL service code
    # EMERGENCY_STOP_COMMON(LIST:=LIST,Bit_Input:=GVL_Alarms.EMERGENCY_STOP_COMMON,ID:=101,Add_Info:='',Group:=App_Error_Group,HushOption:=TRUE);
    Alarmbit_app_temp = row_alm_TAG + "(LIST:=LIST,Bit_Input:=GVL_Alarms." + row_alm_TAG + ",ID:=" + str(row_alm_ID) + ",Add_Info:='',Group:=App_Error_Group,HushOption:=" + str(int(row_alm_HUSH)) + ");"
    Alarms_Monitor_App_STL_List.append(Alarmbit_app_temp)

    # Device.Application.AlarmGroup_IO	1001	ESTOPS_SAFETY_RELAY {CHASSIS.IN0100}	ESTOPS_SAFETY_RELAY {CHASSIS.IN0100}	ESTOPS_SAFETY_RELAY {CHASSIS.IN0100}
    # TABS between! - depends on the number of languages in the system!

    Temp_AlarmGroup_App_Full = "Device.Application.AlarmGroup_App" + "\t" + str(int(row_alm_ID)) + "\t" + row_alm_Default_Text + "\t" + row_alm_Default_Text + "\t" + row_alm_Default_Text
    AlarmGroup_App.append(Temp_AlarmGroup_App_Full)

    print(f'Alarms {row_alm_ID=} {row_alm_TAG=} Done')       


replace_section_in_file(TempFile6, TempFile7,GVL_Alarms_Declare_List,"// GVL_Alarms_Declare")
replace_section_in_file(TempFile7, TempFile8,Alarms_Monitor_App_Declare_List,"// Alarms_Monitor_App_Declare")
replace_section_in_file(TempFile8, TempFile9,Alarms_Monitor_App_STL_List,"// Alarms_Monitor_App_STL_Body")
replace_section_in_file(TempFile9, TempFile10,Alarms_Monitor_IO_Declare_List,"// Alarms_Monitor_IO_Declare")
replace_section_in_file(TempFile10,TempFile11,Alarms_Monitor_IO_STL_List,"// Alarms_Monitor_IO_STL_Body")
replace_section_in_file(TempFile11,TempFile12,HMI_IO_Init_Input_Links_LIST,"// HMI_IO_Init_Input_Links")
replace_section_in_file(TempFile12,TempFile13,HMI_IO_Init_Output_Links_LIST,"// HMI_IO_Init_Output_Links")
replace_section_in_file(TempFile13,TempFile14,HMI_IO_Update_Inputs_LIST,"// HMI_IO_Update_Inputs")
replace_section_in_file(TempFile14,TempFile15,HMI_IO_Update_Outputs_LIST,"// HMI_IO_Update_Outputs")
shutil.copy(TempFile15, EXPORT_Filename) # keep a copy of the excel file we used

with open(EXPORT_Directory + "Device.Application.TextList_IO_Diag.csv", 'w') as TextList_outfile:
    for item in TextList_IO_Diag:
        # write each item on a new line
        TextList_outfile.write("%s\n" % item)
    print(f'Textlist {TextList_outfile.name} Done')
            
with open(EXPORT_Directory + "Device.Application.AlarmGroup_App.csv", 'w') as TextList_outfile:
    for item in AlarmGroup_App:
        # write each item on a new line
        TextList_outfile.write("%s\n" % item)
    print(f'Textlist {TextList_outfile.name} Done')

with open(EXPORT_Directory + "Device.Application.AlarmGroup_IO.csv", 'w') as TextList_outfile:
    for item in AlarmGroup_IO:
        # write each item on a new line
        TextList_outfile.write("%s\n" % item)
    print(f'Textlist {TextList_outfile.name} Done')
