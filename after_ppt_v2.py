JPT.ClearLog()
import csv
import os
import math
import numpy as np
import sys
from pyjdg import *
from pptx.util import Inches, Pt, Emu

def split_and_sort_by_coordinates(big_list):
    big_list_sorted_by_x = sorted(big_list, key=lambda x: x[1])
    group_y_greater_than_x_minus_1 = [sublist for sublist in big_list_sorted_by_x if sublist[2] > big_list_sorted_by_x[0][2] - 0.1]
    group_y_less_than_x_minus_1 = [sublist for sublist in big_list_sorted_by_x if sublist[2] <= big_list_sorted_by_x[0][2] - 0.1]
    group_mid = [sublist for sublist in big_list_sorted_by_x if sublist[2] > big_list_sorted_by_x[0][2] - 0.1 and sublist[2] < big_list_sorted_by_x[0][2] + 0.1]

    if len(group_y_greater_than_x_minus_1) > 20:
        result = [group_y_greater_than_x_minus_1[24], group_y_greater_than_x_minus_1[18], group_y_greater_than_x_minus_1[12],
                  group_y_greater_than_x_minus_1[6], group_y_greater_than_x_minus_1[0],
                  group_y_less_than_x_minus_1[5], group_y_less_than_x_minus_1[11], group_y_less_than_x_minus_1[17]]
    else:
        result = [group_mid[-2], group_mid[1]]

    return result
def split_list_by_z(big_list):
    list_z_up = []
    list_z_down = []
    for sublist in big_list:
        if sublist[3] > 15:
            list_z_up.append(sublist)
        else:
            list_z_down.append(sublist)
    return [list_z_down, list_z_up]
def center(list_node):
    center_coordinate = [0,0,0] 
    for i in range(0,len(list_node)):
        center_coordinate[0] += float(list_node[i][0])/len(list_node)
        center_coordinate[1] += float(list_node[i][1])/len(list_node)
        center_coordinate[2] += float(list_node[i][2])/len(list_node)
    return center_coordinate
def rotate_list_around_line_parallel_to_Y(big_list, pivot_point, alpha):
    alpha_rad = np.radians(-alpha)
    rotation_matrix = np.array([
        [np.cos(alpha_rad), 0, np.sin(alpha_rad)],
        [0, 1, 0],
        [-np.sin(alpha_rad), 0, np.cos(alpha_rad)]
    ])
    pivot_point = np.array(pivot_point)
    rotated_list = []
    for sublist in big_list:
        point = np.array(sublist[1:4])
        translated_point = point - pivot_point
        rotated_translated_point = np.dot(rotation_matrix, translated_point)
        rotated_point = rotated_translated_point + pivot_point
        rotated_sublist = [sublist[0]] + rotated_point.tolist()
        rotated_list.append(rotated_sublist)
    return rotated_list
def angle_between_line_and_OX(p1, p2):
    p1, p2 = np.array(p1), np.array(p2)
    line_vector = p2 - p1
    x_axis_vector = np.array([1, 0, 0])
    cosine_angle = np.dot(line_vector, x_axis_vector) / (np.linalg.norm(line_vector) * np.linalg.norm(x_axis_vector))
    angle = np.degrees(np.arccos(abs(cosine_angle)))
    return angle
def show_results(process,step):
    JPT.Exec(f'CmdShowPostContour(183:1, {{1, 0, {process}, {step}, Displacement, Translational, 1}}, {{1, 1, 0, 0, 0, 0, 0, 0.000000, 0}}, 0, {{0, 0, 0, 0, , , 0}}, {{0, 0, 0, 0, 0, 0, 0, 0.000000, 0}}, 0, {{0, 0, 0, 0, , , 0}}, {{0, 0, 0, 0, 0, 0, 0, 0.000000, 0}}, 0, 0)')
    JPT.Exec(f'CmdShowPostDeformation(183:1, 1, 0, {process}, {step}, Displacement, index, 0, 0.000000, 0, 0.070000, 0, 0.070000, 0.070000, 0.070000, 0)')
    JPT.Exec('CmdDataPaneDeleteAll()')
def center(list_node):
    center_coordinate = [0,0,0] 
    for i in range(0,len(list_node)):
        center_coordinate[0] += float(list_node[i][0])/len(list_node)
        center_coordinate[1] += float(list_node[i][1])/len(list_node)
        center_coordinate[2] += float(list_node[i][2])/len(list_node)
    return center_coordinate
def line_plane_intersection(plane_points, line_points):
    p1, p2, p3 = plane_points
    l1, l2 = line_points
    v1 = np.array(p2) - np.array(p1)
    v2 = np.array(p3) - np.array(p1)
    normal = np.cross(v1, v2)
    a, b, c = normal
    d = -np.dot(normal, p1)
    direction = np.array(l2) - np.array(l1)
    t = -(np.dot(normal, l1) + d) / np.dot(normal, direction)
    intersection = l1 + t * direction
    return intersection.tolist()
def rot(vector):
    theta_ex = -0.34906585 #-20deg
    theta_in = 0.34906585
    vector_rot = [0,0,0]
    if vector[0] > 0:
        vector_rot[0] = math.cos(theta_ex)*vector[0] + math.sin(theta_ex)*vector[2]
        vector_rot[1] = vector[1]
        vector_rot[2] = math.cos(theta_ex)*vector[2] + math.sin(theta_ex)*vector[0]
    if vector[0] < 0:
        vector_rot[0] = math.cos(theta_in)*vector[0] + math.sin(theta_in)*vector[2]
        vector_rot[1] = vector[1]
        vector_rot[2] = math.cos(theta_in)*vector[2] + math.sin(theta_in)*vector[0]
    return  vector_rot
def print_watch_nodes(id):
    JPT.Exec(f"WatchNode({id})")
def import_csv(csv_file_path):
    displacement = []
    with open(csv_file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            displacement.append([float(row['Pos.(X)'])+float(row['Data(X)'])\
                                    ,float(row['Pos.(Y)'])+float(row['Data(Y)'])\
                                    ,float(row['Pos.(Z)'])+float(row['Data(Z)'])])
    return displacement
def export_value_to_list(file_path, list_data, process, line_number):
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
        while len(lines) < line_number + 1:
            lines.append('[]\n')
        selected_line = eval(lines[line_number].strip())
        index = 0 if process == 0 else 1
        while len(selected_line) < index + 1:
            selected_line.append([])
        selected_line[index] = list_data
        lines[line_number] = str(selected_line) + '\n'
        with open(file_path, 'w') as file:
            file.writelines(lines)
    except SyntaxError as e:
        print("", e)
    except IndexError as e:
        print("", e)
    except Exception as e:
        print("", e)
def get_list_from_line(file_path, line_number):
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
        if line_number >= len(lines) or line_number < 0:
            return None, "Line number out of range."
        selected_line = lines[line_number].strip()
        return eval(selected_line), None
    except SyntaxError as e:
        return None, "Syntax error in line content: " + str(e)
    except Exception as e:
        return None, "An error occurred: " + str(e)
def find_node(csv_file_path):
    divide_x_m = []
    divide_x_p = []
    with open(csv_file_path, newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if float(row[3]) < 0:
                divide_x_m.append([int(row[0]), float(row[3]), float(row[4]), float(row[5])])
            else:
                divide_x_p.append([int(row[0]), float(row[3]), float(row[4]), float(row[5])])

    list_16 = []

    divide_x_m_sorted = sorted(divide_x_m, key=lambda x: x[1])
    divide_x_p_sorted = sorted(divide_x_p, key=lambda x: x[1])
    sorted_data_y_m = sorted(divide_x_m_sorted, key=lambda y: y[2])
    sorted_data_y_p = sorted(divide_x_p_sorted, key=lambda y: y[2])

    point_1 = [divide_x_p_sorted[0], divide_x_m_sorted[-1]]
    point_1[0][2] = 0
    point_1[1][2] = 0
    
    for i in range(0,8):    
        list_16.append(sorted_data_y_p[int(len(sorted_data_y_p)*i/8):int(len(sorted_data_y_p)*(i+1)/8)])
    for i in range(0,8):
        list_16.append(sorted_data_y_m[int(len(sorted_data_y_m)*i/8):int(len(sorted_data_y_m)*(i+1)/8)])
    list_16_sort_z = []
    for list in list_16:
        list_16_sort_z.append(sorted(list, key=lambda z: z[3]))

    point_2 = [list_16_sort_z[0][0], list_16_sort_z[-1][0]]
    point_2[0][2] = 0
    point_2[1][2] = 0

    angle = [angle_between_line_and_OX(point_1[0][1:4], point_2[0][1:4]), angle_between_line_and_OX(point_1[1][1:4], point_2[1][1:4])]
    point = [center([point_1[0][1:4],point_2[0][1:4]]),center([point_1[1][1:4],point_2[1][1:4]])]

    zero =0
    list_valve_guide_selected = []
    for list in list_16_sort_z:
        if zero < 8:
            list_valve_and_guide = sorted(rotate_list_around_line_parallel_to_Y(list,point[0], angle[0]), key=lambda z: z[3])
            divide_list = split_list_by_z(list_valve_and_guide)
            list_valve_48 = divide_list[0][48:96]
            list_guide_up = divide_list[1][-32:len(divide_list[1])]
            list_guide_down = divide_list[1][0:32]
            list_valve_guide_selected.append(split_and_sort_by_coordinates(list_valve_48)+split_and_sort_by_coordinates(list_guide_down)+split_and_sort_by_coordinates(list_guide_up))
        else:
            list_valve_and_guide = sorted(rotate_list_around_line_parallel_to_Y(list,point[1], -1*angle[1]), key=lambda z: z[3])
            divide_list = split_list_by_z(list_valve_and_guide)
            list_valve_48 = divide_list[0][48:96]
            list_guide_up = divide_list[1][-32:len(divide_list[1])]
            list_guide_down = divide_list[1][0:32]
            list_valve_guide_selected.append(split_and_sort_by_coordinates(list_valve_48)+split_and_sort_by_coordinates(list_guide_down)+split_and_sort_by_coordinates(list_guide_up))
        ids = [str(sublist[0]) for sublist in list_valve_guide_selected[zero]]
        ids_str = " ".join(ids)
        
        result_str = f'ViewFindEntities("{ids_str}", "Node", 0)'
        JPT.Exec(result_str)
        zero += 1
    a = sorted(rotate_list_around_line_parallel_to_Y(list_16_sort_z[11],point[1], -1*angle[1]), key=lambda z: z[3])
    c = split_list_by_z(a)
    list_valve_48 = c[0][0:96]
    list_guide_up = c[1][-32:len(divide_list[1])]
    list_guide_down = c[1][0:32]
    b = []
    for i in range(len(list_valve_48)):
        b.append(list_valve_48[i][0])
    print(b)
    e = list_16_sort_z[11]
    d = []
    for i in range(len(e)):
        d.append(e[i][0])
    print(d)
    
    print(len(list_16_sort_z[8]))
    print(len(list_16_sort_z[9]))
    print(len(list_16_sort_z[10]))
    print(len(list_16_sort_z[11]))
def insert_results_to_excel(file_excel,process,step):
    show_results(process,step)
    directory = os.path.dirname(file_excel)
    csv_file_path = os.path.join(directory, "192_id.csv")
    id_list = []
    with open(csv_file_path, newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            id_list.append(int(row[0]))
    id_list_add = ["10:"+str(value) for value in id_list]
    formatted_id = ", ".join(id_list_add)
    ids = f"[{formatted_id}]"
    print_watch_nodes(ids)
    file_csv = os.path.join(directory, "data.csv")
    JPT.Exec(f'CmdDataPaneSaveToFile("{file_csv}")')
    displacement = import_csv(file_csv)
    if displacement[0][0] < 0:
        displacement = displacement[::-1]
    list_denta_ex = []
    list_denta_in = []
    list_d_ex = []
    list_d_in = []
    for i in range(0,8):
        group_values_ex = displacement[12*i:8+12*i]
        group_guide_ex  = displacement[8+12*i:12+12*i]
        group_values_in = displacement[96+12*i:8+96+12*i]
        group_guide_in  = displacement[8+96+12*i:12+96+12*i]
        center_guide_ex = [center([group_guide_ex[0],group_guide_ex[1]]),center([group_guide_ex[2],group_guide_ex[3]])]
        center_guide_in = [center([group_guide_in[0],group_guide_in[1]]),center([group_guide_in[2],group_guide_in[3]])]
        intersection_point_ex = line_plane_intersection([group_values_ex[0],group_values_ex[3],group_values_ex[5]], center_guide_ex)
        intersection_point_in = line_plane_intersection([group_values_in[0],group_values_in[3],group_values_in[5]], center_guide_in)
        center_valve_ex_8_point = center(group_values_ex)
        center_valve_in_8_point = center(group_values_in)
        intersection_point_ex_rot = rot(intersection_point_ex)
        intersection_point_in_rot = rot(intersection_point_in)
        center_valve_ex_8_point_rot = rot(center_valve_ex_8_point)
        center_valve_in_8_point_rot = rot(center_valve_in_8_point)
        denta_ex = [round(2000 * (b_i - a_i), 1) for a_i, b_i in zip(intersection_point_ex_rot, center_valve_ex_8_point_rot)]
        denta_in = [round(2000 * (b_i - a_i), 1) for a_i, b_i in zip(intersection_point_in_rot, center_valve_in_8_point_rot)]
        d_ex = round(math.sqrt(denta_ex[0]**2+denta_ex[1]**2+denta_ex[2]**2),1)
        d_in = round(math.sqrt(denta_in[0]**2+denta_in[1]**2+denta_in[2]**2),1)
        list_denta_ex.append(denta_ex[0])
        list_denta_ex.append(denta_ex[1])
        list_denta_in.append(denta_in[0])
        list_denta_in.append(denta_in[1])
        list_d_ex.append(d_ex)
        list_d_in.append(d_in)
    return list_d_ex, list_d_in, list_denta_ex, list_denta_in
def onGetButton1Clicked(dlg):
    file_excel = dlg.get_item_text(name="Browser1")
    directory = os.path.dirname(file_excel)
    csv_file_path = os.path.join(directory, "all_node.csv")
    # JPT.Exec(f'CmdDataPaneSaveToFile("{csv_file_path}")')
    JPT.ClearAllSelection()
    find_node(csv_file_path)
def onGetButton2Clicked(dlg):
    file_excel = dlg.get_item_text(name="Browser1")
    process=dlg.get_item_text(name="ComboBox1")
    step=dlg.get_item_text(name="ComboBox2")
    print(type(process))
    d_ex, d_in, denta_ex, denta_in = insert_results_to_excel(file_excel,process,step)
    directory = os.path.dirname(file_excel)
    file_path = os.path.join(directory, "New Text Document.txt")
    if process == "0":
        export_value_to_list(file_path, d_ex, 0, 0)
        export_value_to_list(file_path, d_in, 0, 1)
        export_value_to_list(file_path, denta_ex, 0, 2)
        export_value_to_list(file_path, denta_in, 0, 3)
    elif process == "4":
        export_value_to_list(file_path, d_ex, 1, 0)
        export_value_to_list(file_path, d_in, 1, 1)
        export_value_to_list(file_path, denta_ex, 1, 2)
        export_value_to_list(file_path, denta_in, 1, 3)

def onGetButton3Clicked(dlg):
    file_excel = dlg.get_item_text(name="Browser1")
    directory = os.path.dirname(file_excel)
    sys.path.append(directory)
    file_path = os.path.join(directory, "New Text Document.txt")
    list_data_d_ex, error = get_list_from_line(file_path, 0)
    list_data_d_in, error = get_list_from_line(file_path, 1)
    list_data_denta_ex, error = get_list_from_line(file_path, 2)
    list_data_denta_in, error = get_list_from_line(file_path, 3)
    import chart_line_function
    import chart_column_function
    import insert_table_function
    pptx = os.path.join(directory, "sample.pptx")
    slide_insert = 1
    chart_line_function.chart_line(pptx, list_data_denta_ex, Inches(2.5), Inches(5.5), Inches(2.1), Inches(2.1), slide_insert)
    chart_line_function.chart_line(pptx, list_data_denta_in, Inches(6.8), Inches(5.5), Inches(2.1), Inches(2.1), slide_insert)

    chart_column_function.chart_column(pptx, list_data_d_ex[0], list_data_d_ex[1], Inches(0.4), Inches(2.2), Inches(3.8), Inches(2.5), slide_insert)
    chart_column_function.chart_column(pptx, list_data_d_in[0], list_data_d_in[1], Inches(5.5), Inches(2.2), Inches(3.8), Inches(2.5), slide_insert)

    insert_table_function.insert_table(pptx,slide_insert,1,2,list_data_d_ex[0],9,9)
    insert_table_function.insert_table(pptx,slide_insert,2,2,list_data_d_in[0],9,9)
    insert_table_function.insert_table(pptx,slide_insert,3,3,list_data_denta_ex[0],7,17)
    insert_table_function.insert_table(pptx,slide_insert,4,3,list_data_denta_in[0],7,17)

    insert_table_function.insert_table(pptx,slide_insert,1,3,list_data_d_ex[1],9,9)
    insert_table_function.insert_table(pptx,slide_insert,2,3,list_data_d_in[1],9,9)
    insert_table_function.insert_table(pptx,slide_insert,3,4,list_data_denta_ex[1],7,17)
    insert_table_function.insert_table(pptx,slide_insert,4,4,list_data_denta_in[1],7,17)
def main():
    dlg=JDGCreator(title="Valve Misalignment")
    dlg.add_node_selector()
    dlg.add_groupbox(name="GroupBox1",text="Data",layout="Window")
    dlg.add_label(name="Label1",text="Excel Link",text_halign="left",text_valign="top",layout="GroupBox1")
    dlg.add_browser(name="Browser1",mode="file",file_filter="All Files(*.*)|*.*|||*.*|||*.*|||*.*|||*.*||",layout="GroupBox1")
    dlg.add_button(name="Button1",text="Get Node",width=60,height=22,bk_color=15790320,layout="GroupBox1")
    dlg.add_groupbox(name="GroupBox2",text="Process",layout="Window")
    dlg.set_groupbox_orientation(name="GroupBox2",orientation="horizontal")
    dlg.add_label(name="Label2",text="Process",text_halign="left",text_valign="top",layout="GroupBox2")
    dlg.add_combobox(name="ComboBox1",options=["0","4"],index=0,layout="GroupBox2")
    dlg.add_label(name="Label3",text="Step",text_halign="left",text_valign="top",layout="GroupBox2")
    dlg.add_combobox(name="ComboBox2",options=["0","9"],index=0,layout="GroupBox2")
    dlg.add_groupbox(name="GroupBox3",text="Make Report",layout="Window")
    dlg.add_button(name="Button2",text="Insert to PPT",width=100,height=22,bk_color=15790320,layout="GroupBox3")
    dlg.generate_window()

    dlg.on_button_clicked(name="Button1",callfunc=onGetButton1Clicked)
    dlg.on_button_clicked(name="Button2",callfunc=onGetButton3Clicked)
    dlg.on_dlg_apply(callfunc=onGetButton2Clicked)
    dlg.on_dlg_ok(callfunc=onGetButton2Clicked)
if __name__=='__main__':
    main()