import os    
import json    
import pandas as pd    
import matplotlib.pyplot as plt    
from tkinter import Tk, Button, Entry, filedialog, messagebox    
import tkinter as tk    
def select_file():    
    file_path = filedialog.askdirectory()    
    if file_path:    
        entry_path.delete(0, 'end')    
        entry_path.insert(0, file_path)    
    
def process_files():    
    directory = entry_path.get()    
    if not os.path.isdir(directory):    
        messagebox.showerror("错误", "请选择一个有效的文件夹")    
        return    
        
    # Create a Pandas Excel writer using XlsxWriter as the engine    
    excel_filename = os.path.join(directory, 'dvh.xlsx')    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:    
        for filename in os.listdir(directory):    
            if filename.endswith('.json'):    
                json_path = os.path.join(directory, filename)    
                with open(json_path, 'r', encoding='utf-8') as f:    
                    data = json.load(f)    
                    roi_data = extract_data(data, filename)    #从这里开始出发解析
                        
                    # Create DataFrame and save to a new sheet    
                    df = pd.DataFrame(roi_data)    
                    sheet_name = os.path.splitext(filename)[0]  # Use the filename without extension as sheet name    
                    df.to_excel(writer, sheet_name=sheet_name, index=False)    
                        
                    # Adjust column widths for the current sheet    
                    worksheet = writer.sheets[sheet_name]    
                    for column in worksheet.columns:    
                        max_length = 0    
                        column = [cell for cell in column]    
                        for cell in column:    
                            try:    
                                if len(str(cell.value)) > max_length:    
                                    max_length = len(str(cell.value))    
                            except:    
                                pass    
                        adjusted_width = (max_length + 2)    
                        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width    
    messagebox.showinfo("成功","所有文件处理成功！")
    
def extract_data(data, filename):    
    roi_data = []    
    for roi in data['DvhData']:    
        roi_name = roi['RoiName']    
        absolute_volume = roi['RoiStatisticsBaseInfo']['absoluteVolume'].replace(' cm3', '')  # Remove unit            
        max_dose = roi['RoiStatisticsDoseInfo']['maxDose'].replace(' cGy', '')  # Remove unit    
        min_dose = roi['RoiStatisticsDoseInfo']['minDose'].replace(' cGy', '')  # Remove unit    
        mean_dose = roi['RoiStatisticsDoseInfo']['meanDose'].replace(' cGy', '')  # Remove unit    
        standard_deviation = roi['RoiStatisticsDoseInfo']['standardDeviation']
            
        # Extract doseByVolumeList        
        dose_by_volume = roi['RoiStatisticsDoseInfo']['doseByVolumeList']        
        # dose_by_volume_dict = {f'D{dv["relativeVolume"].replace("%", "")}': dv['dose'] for dv in dose_by_volume} 
        # dose_by_volume_dict = {f'D{dv["relativeVolume"].replace("%", "")}': dv['dose'].replace(' cGy', '') for dv in dose_by_volume}       
        dose_by_volume_dict = {f'D{dv["relativeVolume"].replace("%", "")} (cGy)': dv['dose'].replace(' cGy', '') for dv in dose_by_volume}           
                
        # Extract volumeByDoseList        
        volume_by_dose = roi['RoiStatisticsDoseInfo']['volumeByDoseList']        
        volume_by_dose_dict = {}    
        for vd in volume_by_dose:    
            dose_value = vd['dose'].replace(' cGy', '')  # Remove the unit               
            relative_volume = vd.get('relativeVolume')  # Get relativeVolume if exists        
            absolute_volume_from_vd = vd.get('absoluteVolume')  # Get absoluteVolume if exists      ################注意跟前面的absoluteVolume区分########## 
                    
            if relative_volume is not None:  # If relativeVolume exists        
                volume_by_dose_dict[f'V{str(int(float(dose_value)))}(%)'] = relative_volume.replace('%', '')           
            elif absolute_volume_from_vd is not None:  # If absoluteVolume exists        
                # volume_by_dose_dict[f'V{str(int(float(dose_value)))}(cm3)'] = absolute_volume.replace(' cm3', '')
                 # Only update if absolute_volume is not already set    
                if absolute_volume is None or absolute_volume == '':    
                    absolute_volume = absolute_volume_from_vd.replace(' cm3', '')  # Update only if it's empty     
        # Combine all data    
        roi_info = {    
            'RoiName': roi_name,    
            'absoluteVolume': absolute_volume,    
            'maxDose': max_dose,    
            'minDose': min_dose,    
            'meanDose': mean_dose,    
            'standardDeviation': standard_deviation,    
            **dose_by_volume_dict,    
            **volume_by_dose_dict    
        }    
        roi_data.append(roi_info)    
            
        # Plot DVH    
        # plot_dvh(roi, filename)    #控制画图开关############################################################################################################################
    
    return roi_data    
    
def plot_dvh(roi, filename):    
    dvh_list = roi['DvhList']    
    doses = [d['dose'] for d in dvh_list]    
    relative_volumes = [float(d['relativeVolume'].replace('%', '')) for d in dvh_list]  # Remove '%' and convert to float    
    
    plt.figure()    
    plt.plot(doses, relative_volumes)    
    plt.xlabel('Dose (cGy)')    
    plt.ylabel('Relative Volume (%)')    
    plt.title(f'DVH Curve for {roi["RoiName"]}')    
    plt.grid()    
        
    # Save the plot    
    plot_filename = os.path.splitext(filename)[0] + '.png'    
    plt.savefig(os.path.join(entry_path.get(), plot_filename))    
    plt.close()    
    
# Create the GUI      
def center_window(window, width, height):    
    # 获取屏幕的宽度和高度    
    screen_width = window.winfo_screenwidth()    
    screen_height = window.winfo_screenheight()     
    # 计算窗口的 x 和 y 坐标    
    x = (screen_width // 2) - (width // 2)    
    y = (screen_height // 2) - (height // 2)    
    # 设置窗口的位置    
    window.geometry(f'{width}x{height}+{x}+{y}')
root = Tk()    
root.title("JSON to Excel")  
  
entry_path = Entry(root, width=50)    
# entry_path.pack(pady=10)    
window_width = 400    
window_height = 150      
# 调用函数使窗口居中    
center_window(root, window_width, window_height)     
button_select = Button(root, text="选择文件夹", command=select_file)    
# button_select.pack(pady=5)   
    
button_process = Button(root, text="确定", command=process_files)    
# button_process.pack(pady=20)    

label = tk.Label(root,text="先选择json文件的路径，再点确定", font=("Trail", 14))    
# label.pack(pady=10)  # 使用 pack 方法添加到窗口中，并设置上下边距    

entry_path.place(x=10,y=10,)
button_select.place(x=10,y=50)    
button_process.place(x=100,y=50)  
label.place(x=10,y=100) 
root.mainloop()    
