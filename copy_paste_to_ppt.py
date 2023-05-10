# Grab the Active Instance of Excel.
import win32com.client 


def copy_paste_to_PPT(list_folder,list_folder_surface,path_in_folder,ppt_path):
    
    for i in list_folder:
        
        path_to_plot_pressure = path_in_folder + i + '\\0.1\\Pressure_' + i + '.xlsx'
        
        ExcelApp = win32com.client.Dispatch("Excel.Application")
        ExcelApp.Visible = True
        path_to_plot_velocity = path_in_folder + i + '\\0.1\\Velocity_magnitude_' + i + '.xlsx'
        # Grab the workbook with the charts.
        xlWorkbook = ExcelApp.Workbooks.Open(path_to_plot_pressure)
        xlWorkbook_1 = ExcelApp.Workbooks.Open(path_to_plot_velocity)

        # Create a new instance of PowerPoint and make sure it's visible.
        PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        PPTApp.Visible = True

        # Add a presentation to the PowerPoint Application, returns a Presentation Object.
        PPTPresentation = PPTApp.Presentations.Open(ppt_path, ReadOnly= False)

        # Loop through each Worksheet.
        for xlWorksheet in xlWorkbook.Worksheets:

            # Grab the ChartObjects Collection for each sheet.
            xlCharts = xlWorksheet.ChartObjects()
            
            # Loop through each Chart in the ChartObjects Collection.
            for index, xlChart in enumerate(xlCharts):
                # Each chart needs to be on it's own slide, so at this point create a new slide.
                PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=11)  # 12 is a blank layout
                # Copy the chart.
                xlChart.Copy()

                # Paste the Object to the Slide
                PPTSlide.Shapes.PasteSpecial(DataType=1)
            
            # Save the presentation.
        for xlWorksheet in xlWorkbook_1.Worksheets:

            # Grab the ChartObjects Collection for each sheet.
            xlCharts = xlWorksheet.ChartObjects()
            
            # Loop through each Chart in the ChartObjects Collection.
            for index, xlChart in enumerate(xlCharts):
                # Each chart needs to be on it's own slide, so at this point create a new slide.
                PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=11)  # 12 is a blank layout
                # Copy the chart.
                xlChart.Copy()

                # Paste the Object to the Slide
                PPTSlide.Shapes.PasteSpecial(DataType=1)
            
            # Save the presentation.

        PPTPresentation.SaveAs(ppt_path)
        PPTPresentation.Close()
        xlWorkbook.Close()
        xlWorkbook_1.Close()


    for i in list_folder_surface:

        path_to_plot_pressure_surface = path_in_folder + i + "\\0\\Maximum_Pressure_" + i + ".xlsx"
        ExcelApp = win32com.client.Dispatch("Excel.Application")
        ExcelApp.Visible = True
        path_to_plot_velocity_surface = path_in_folder + i + "\\0\\Maximum_Velocity_" + i + ".xlsx"
        # Grab the workbook with the charts.
        xlWorkbook = ExcelApp.Workbooks.Open(path_to_plot_pressure_surface)
        xlWorkbook_1 = ExcelApp.Workbooks.Open(path_to_plot_velocity_surface)

        # Create a new instance of PowerPoint and make sure it's visible.
        PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        PPTApp.Visible = True

        # Add a presentation to the PowerPoint Application, returns a Presentation Object.
        PPTPresentation = PPTApp.Presentations.Open(ppt_path, ReadOnly= False)

        # Loop through each Worksheet.
        for xlWorksheet in xlWorkbook.Worksheets:

            # Grab the ChartObjects Collection for each sheet.
            xlCharts = xlWorksheet.ChartObjects()
            
            # Loop through each Chart in the ChartObjects Collection.
            for index, xlChart in enumerate(xlCharts):
                # Each chart needs to be on it's own slide, so at this point create a new slide.
                PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=11)  # 12 is a blank layout
                # Copy the chart.
                xlChart.Copy()

                # Paste the Object to the Slide
                PPTSlide.Shapes.PasteSpecial(DataType=1)
            
            # Save the presentation.
        for xlWorksheet in xlWorkbook_1.Worksheets:

            # Grab the ChartObjects Collection for each sheet.
            xlCharts = xlWorksheet.ChartObjects()
            
            # Loop through each Chart in the ChartObjects Collection.
            for index, xlChart in enumerate(xlCharts):
                # Each chart needs to be on it's own slide, so at this point create a new slide.
                PPTSlide = PPTPresentation.Slides.Add(Index=index + 1, Layout=11)  # 12 is a blank layout
                # Copy the chart.
                xlChart.Copy()

                # Paste the Object to the Slide
                PPTSlide.Shapes.PasteSpecial(DataType=1)
            
            # Save the presentation.

        PPTPresentation.SaveAs(ppt_path)
        PPTPresentation.Close()
        xlWorkbook.Close()
        xlWorkbook_1.Close()
