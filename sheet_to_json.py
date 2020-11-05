#!/usr/bin/python
# coding=utf-8
import argparse
import xlrd
import json
import datetime

def insert_agent_ini_status(f,row_num,worksheet):
        f.write('\n\t\t"ref_car": "%s",'% worksheet.cell(row_num,2).value)
        f.write('\n\t\t"w": %s,'% worksheet.cell(row_num + 1,2).value)
        f.write('\n\t\t"l": %s,'% worksheet.cell(row_num + 2,2).value)
        f.write('\n\t\t"h": %s,'% worksheet.cell(row_num + 3,2).value)
        f.write('\n\t\t"yaw": %s,'% worksheet.cell(row_num + 4,2).value)
        f.write('\n\t\t"ds": %s,'% worksheet.cell(row_num + 5,2).value)
        f.write('\n\t\t"dl": %s,'% worksheet.cell(row_num + 6,2).value)
        f.write('\n\t\t"dv": %s,'% worksheet.cell(row_num + 7,2).value)
        f.write('\n\t\t"da": %s,'% worksheet.cell(row_num + 8,2).value)
        f.write('\n\t\t"vmx": %s,'% worksheet.cell(row_num + 9,2).value)
        f.write('\n\t\t"vmn": %s,'% worksheet.cell(row_num + 10,2).value)
        return

def insert_agent_actions(f,row_num,action_num_obj,worksheet):
        f.write('\n\t\t"actions": [')
        print "action_num_obj =", action_num_obj
        action_num = 0
        while action_num < action_num_obj:
                action_num += 1
                current_row = row_num + action_num
                action_type = worksheet.cell(current_row,3).value
                f.write("\n\t\t{")
                #f.write('\n\t\t\t"action": action_type,')
                if worksheet.cell(current_row,3).value == "ta":
                        f.write('\n\t\t\t"action": "' + action_type + '",')
                        f.write('\n\t\t\t"expr": "x{self} - x{ego} <= {{dist}}",')
                        f.write('\n\t\t\t"params": {')
                        f.write('\n\t\t\t\t"dist": %s'% worksheet.cell(current_row,6).value)
                        f.write('\n\t\t\t}')
                elif worksheet.cell(current_row,3).value in ['lf','lc','ci','co','rob','rtc','lm','rv','cm']:
                                f.write('\n\t\t\t"action": "' + action_type + '",')
                                f.write('\n\t\t\t"params": {')
                                f.write('\n\t\t\t\t"t": %s,'% worksheet.cell(current_row,4).value)
                                if worksheet.cell(current_row,3).value in ['lf','ci','rtc','rv','cm']:
                                        f.write('\n\t\t\t\t"ax": %s'% worksheet.cell(current_row,5).value)
                                elif worksheet.cell(current_row,3).value in ['lc','co','rob','lm']:
                                        f.write('\n\t\t\t\t"ax": %s,'% worksheet.cell(current_row,5).value)
                                        if worksheet.cell(current_row,3).value == "lm":
                                                f.write('\n\t\t\t\t"dl": %s'% worksheet.cell(current_row,6).value)
                                        elif worksheet.cell(current_row,3).value == "rob":
                                                f.write('\n\t\t\t\t"dl": %s,'% worksheet.cell(current_row,6).value)
                                                f.write('\n\t\t\t\t"left": %s'% worksheet.cell(current_row,7).value)
                                        elif worksheet.cell(current_row,3).value in ['lc','co']:
                                                f.write('\n\t\t\t\t"left": %s'% worksheet.cell(current_row,7).value)
                                f.write('\n\t\t\t}')
                elif worksheet.cell(current_row,3).value in ['tlc','tco']:
                        f.write('\n\t\t\t"action": "' + action_type + '",')
                        f.write('\n\t\t\t"params": {')
                        f.write('\n\t\t\t\t"left": %s'% worksheet.cell(current_row,7).value)
                        f.write('\n\t\t\t}')
                elif worksheet.cell(current_row,3).value == "tci":
                        f.write('\n\t\t\t"action": "' + action_type + '",')
                        f.write('\n\t\t\t"params": {')
                        f.write('\n\t\t\t}')
                else:
                         print "No such action type, please confirm your cmd"
                         print "action_num_final2=", action_num
                         print "action_type is ", worksheet.cell(current_row,3).value
                
                if action_num < action_num_obj:
                        f.write("\n\t\t},")
                else:        
                        f.write('\n\t\t}')
        #actions ]
        f.write('\n\t\t]')
        return

def xlsx_to_json(worksheet,f):
        f.write("{")
        #scenario_name_prefix
        f.write('\n\t"prefix": "%s",'% worksheet.cell(1,1).value)
        f.write('\n\t"description": "scenario name could be Prefix_id.prototxt rating from 00001, for example, acc_cut_in_1_00001.prototxt",')
        #maps
        f.write('\n\t"maps": [')
        #map1
        f.write("\n\t{")
        f.write('\n\t\t"lane_width": 3.75,')
        f.write('\n\t\t"left_lanes": 1,')
        f.write('\n\t\t"right_lanes": 1,')
        f.write('\n\t\t"builders": [')
        f.write('\n\t\t\t"line:dist=1000"')
        f.write('\n\t\t]')
        f.write("\n\t},")
        #map2
        f.write("\n\t{")
        f.write('\n\t\t"lane_width": 3.75,')
        f.write('\n\t\t"left_lanes": 1,')
        f.write('\n\t\t"right_lanes": 1,')
        f.write('\n\t\t"builders": [')
        f.write('\n\t\t\t"line:dist=100",')
        f.write('\n\t\t\t"spiral:dist=100",')
        f.write('\n\t\t\t"cone:dist=800,radius=500,left=False"')
        f.write('\n\t\t]')
        f.write("\n\t},")
        #map3
        f.write("\n\t{")
        f.write('\n\t\t"lane_width": 3.75,')
        f.write('\n\t\t"left_lanes": 1,')
        f.write('\n\t\t"right_lanes": 1,')
        f.write('\n\t\t"builders": [')
        f.write('\n\t\t\t"line:dist=100",')
        f.write('\n\t\t\t"spiral:dist=100",')
        f.write('\n\t\t\t"cone:dist=800,radius=500,left=True"')
        f.write('\n\t\t]')
        f.write("\n\t}")
        f.write('\n\t],')
        #ego info
        f.write('\n\t"ego":')
        f.write("\n\t{")
        f.write('\n\t\t"w": %s,'% worksheet.cell(5,2).value)
        f.write('\n\t\t"l": %s,'% worksheet.cell(6,2).value)
        f.write('\n\t\t"h": %s,'% worksheet.cell(7,2).value)
        f.write('\n\t\t"v": %s,'% worksheet.cell(8,2).value)
        f.write('\n\t\t"a": %s,'% worksheet.cell(9,2).value)
        f.write('\n\t\t"yaw": %s'% worksheet.cell(10,2).value)
        f.write("\n\t},")
        #objects
        f.write('\n\t"agents": [')
        #object 1 start {
        f.write("\n\t{")
        #insert initilize parameters for object 1
        insert_agent_ini_status(f,11,worksheet)
        action_num_obj1 = worksheet.cell(22,3).value
        print "action_num_obj1=", str(action_num_obj1)
        #insert action lists for object 1
        insert_agent_actions(f,22,action_num_obj1,worksheet)
        #object 1 finish }
        f.write("\n\t}")
        #f.write("\n\t},")
        #object 2 start {
        #f.write("\n\t{")
        #insert initilize parameters for object 2
        #insert_agent_ini_status(f,36,worksheet)
        #action_num_obj2 = worksheet.cell(47,3).value
        #insert action lists for object 2
        #insert_agent_actions(f,47,action_num_obj2,worksheet)
        #object 2 finish }
        #f.write("\n\t}")
        #objects finish ]
        f.write('\n\t]')
        #the whole file finish }
        f.write("\n}")

        return
        

def main():
        parser = argparse.ArgumentParser(description="convert .xlsx to json file")
        parser.add_argument('-w','--workbook', help='.xlsx to read')
        parser.add_argument('-s','--sheet', help='sheet on .xlsx to read')
        args = parser.parse_args()

        wb = xlrd.open_workbook(args.workbook)
        work_sheet = wb.sheet_by_name(args.sheet)
        path = work_sheet.name.lower() + ".json"
        json_file = open(path,"w+") 
        xlsx_to_json(work_sheet,json_file)


if __name__ == "__main__":
        main()
