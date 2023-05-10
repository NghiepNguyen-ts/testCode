import xlsxwriter


def file_processing(list_folder,list_folder_surface,path_in_folder):
    
    for name in list_folder:
        path_to_read_pressure = path_in_folder + name + "/0.1/static(p)"
        path_to_write_pressure = path_in_folder + name + "/0.1/static(p)_edit"
        path_to_plot_pressure = path_in_folder + name + "/0.1/Pressure_" + name + ".xlsx"
        path_to_read_velocity = path_in_folder + name + "/0.1/mag(U)"
        path_to_write_velocity = path_in_folder + name + "/0.1/U_edit"
        path_to_plot_velocity = path_in_folder + name + "/0.1/Velocity_magnitude_" + name + ".xlsx"
        ##### Read static pressure 
        file_edit = list()
        u_edit = list()
        x = list()
        y = list()
        p_deduct = list()
        initial_pressure = float(0.4e6)

        with open(path_to_read_pressure,'r') as fr:
            a = fr.readlines()
            flat = '#'
            for i in range(len(a)):
                if not a[i].startswith(flat):
                    file_edit.append(a[i])      
        with open(path_to_write_pressure,'w') as fw:
            for i in file_edit:
                fw.writelines(i)
        with open(path_to_write_pressure,'r') as fr:
            a = fr.readlines()
            for i in range(len(a)):
                tmp = a[i].split()
                x.append(tmp[0])
                y.append(tmp[1])
        for i in range(len(y)):
            tmp = float(y[i]) - initial_pressure
            p_deduct.append(tmp)

        ### Read velocity and export velocity magnitude
        t = list()
        u_magnitude = list()
        with open(path_to_read_velocity,'r') as fr:
            a = fr.readlines()
            flat = '#'
            for i in range(len(a)):
                if not a[i].startswith(flat):
                    u_edit.append(a[i]) 
        with open(path_to_write_velocity,'w') as fw:
            for i in u_edit:
                fw.writelines(i)    
        with open(path_to_write_velocity,'r') as fr:
            a = fr.readlines()
            for i in range(len(a)):
                tmp = a[i].split()
                t.append(tmp[0])
                u_magnitude.append(tmp[1]) 

        ## Plot pressure in excel
        workbook = xlsxwriter.Workbook(path_to_plot_pressure)
        worksheet = workbook.add_worksheet()
        for i in range(len(x)):
            worksheet.write(i, 0, x[i])
            worksheet.write(i, 1, p_deduct[i])
        chart = workbook.add_chart({'type': 'scatter',
                                    'subtype': 'smooth'} )

        # Configure the chart. In simplest case we add one or more data series.
        chart.add_series({'categories': '=Sheet1!$A$1:$A$4000','values': '=Sheet1!$B$1:$B$4000'})
        chart.set_legend({'none': True}) 
        chart.set_title({
        'name': 'Pressure - ' + str(name) })
        # Add x-axis label 
        chart.set_size({'width': 600, 'height':400})
        chart.set_x_axis({'name': 'Time(s/10)',
                        'label_position': 'low',
                        'min': 0, 'max': 4000})   
        # Add y-axis label 
        chart.set_y_axis({'name': 'Pressure (Pa)'}) 
        worksheet.insert_chart('A7', chart)
        workbook.close()

        ## Plot velocity in excel
        workbook = xlsxwriter.Workbook(path_to_plot_velocity)
        worksheet = workbook.add_worksheet()
        for i in range(len(t)):
            worksheet.write(i, 0, t[i])
            worksheet.write(i, 1, float(u_magnitude[i]))
        chart = workbook.add_chart({'type': 'scatter',
                                    'subtype': 'smooth'} )

        # Configure the chart. In simplest case we add one or more data series.
        chart.add_series({'categories': '=Sheet1!$A$1:$A$4000','values': '=Sheet1!$B$1:$B$4000'})
        chart.set_legend({'none': True}) 
        chart.set_title({
        'name': 'Velocity - ' + str(name) })
        # Add x-axis label 
        chart.set_size({'width': 600, 'height':400})
        chart.set_x_axis({'name': 'Time(s/10)',
                        'label_position': 'low',
                        'min': 0, 'max': 4000})   
        # Add y-axis label 
        chart.set_y_axis({'name': 'Velocity (m/s)' }) 
        worksheet.insert_chart('A7', chart)
        workbook.close()

    for name in list_folder_surface:
        path_to_read_surface = path_in_folder + name + "\\0\\surfaceFieldValue.dat" 
        path_to_write_surface = path_in_folder + name + "\\0\\surfaceFieldValue_edit.dat"
        path_to_plot_pressure_surface = path_in_folder + name + "\\0\\Maximum_Pressure_" + name  +".xlsx"
        path_to_plot_velocity_surface = path_in_folder + name + "\\0\\Maximum_Velocity_" + name + ".xlsx"
        file_edit = list()
        t_surface = list()
        p_surface = list()
        u_magnitude_surface = list()    
        p_surface_deduct = list()
        initial_pressure = float(0.4e6)
        with open(path_to_read_surface,'r') as fr:
            a = fr.readlines()
            flat = '#'
            for i in range(len(a)):
                if not a[i].startswith(flat):
                    file_edit.append(a[i])
        with open(path_to_write_surface,'w') as fw:
            for i in file_edit:
                fw.writelines(i)

        with open(path_to_write_surface,'r') as fR:
            a = fR.readlines()
            for i in range(len(a)):
                tmp = a[i].split()
                t_surface.append(tmp[0])
                u_magnitude_surface.append(tmp[1])
                p_surface.append(tmp[3])
        for i in range(len(p_surface)):
            tmp = float(p_surface[i]) - initial_pressure
            p_surface_deduct.append(tmp)
        ## Plot pressure surface in excel
        workbook = xlsxwriter.Workbook(path_to_plot_pressure_surface)
        worksheet = workbook.add_worksheet()
        for i in range(len(t_surface)):
            worksheet.write(i, 0, t_surface[i])
            worksheet.write(i, 1, p_surface_deduct[i])
        chart = workbook.add_chart({'type': 'scatter',
                                    'subtype': 'smooth'} )

        # Configure the chart. In simplest case we add one or more data series.
        chart.add_series({'categories': '=Sheet1!$A$1:$A$4000','values': '=Sheet1!$B$1:$B$4000'})
        chart.set_legend({'none': True}) 
        chart.set_title({
        'name': 'Maximum Pressure - ' + str(name) })
        # Add x-axis label 
        chart.set_size({'width': 600, 'height':400})
        chart.set_x_axis({'name': 'Time (s/10)',
                        'label_position': 'low',
                        'min': 0, 'max': 4000})   
        # Add y-axis label 
        chart.set_y_axis({'name': 'Pressure (Pa)'}) 
        worksheet.insert_chart('A7', chart)
        workbook.close()
    ## Plot velocity in excel
        workbook = xlsxwriter.Workbook(path_to_plot_velocity_surface)
        worksheet = workbook.add_worksheet()
        for i in range(len(t_surface)):
            worksheet.write(i, 0, t_surface[i])
            worksheet.write(i, 1, float(u_magnitude_surface[i]))
        chart = workbook.add_chart({'type': 'scatter',
                                    'subtype': 'smooth'} )

        # Configure the chart. In simplest case we add one or more data series.
        chart.add_series({'categories': '=Sheet1!$A$1:$A$4000','values': '=Sheet1!$B$1:$B$4000'})
        chart.set_legend({'none': True}) 
        chart.set_title({
        'name': 'Maximum Velocity - ' + str(name) })
        # Add x-axis label 
        chart.set_size({'width': 600, 'height':400})
        chart.set_x_axis({'name': 'Time (s/10)',
                        'label_position': 'low',
                        'min': 0, 'max': 4000})   
        # Add y-axis label 
        chart.set_y_axis({'name': 'Velocity (m/s)' }) 
        worksheet.insert_chart('A7', chart)
        workbook.close()