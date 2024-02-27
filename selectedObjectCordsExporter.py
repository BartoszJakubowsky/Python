from qgis.core import QgsProject
from PyQt5.QtWidgets import QFileDialog
from qgis.utils import iface
import openpyxl

    
    
# Get the active layer
layer = iface.activeLayer()

# Check if a layer is selected
if layer is not None:
    # Get the selected features
    selected_features = layer.selectedFeatures()

    # Check if any features are selected
    if selected_features:
        # Prompt the user to choose a location to save the Excel file
        excel_path, _ = QFileDialog.getSaveFileName(None, 'Save Excel File', '', 'Excel Files (*.xlsx)')

        # Check if the user selected a location
        if excel_path:
            # Create an Excel workbook and add a worksheet
            workbook = openpyxl.Workbook()
            worksheet = workbook.active

            # Write header row
            worksheet.append(['X', 'Y'])

            # Write coordinates of selected features
            for feature in selected_features:
                geom = feature.geometry()
                
                if geom.type() == 0:  # Point geometry
                    worksheet.append([geom.asPoint().y(), geom.asPoint().x()])
                elif geom.isMultipart():
                    for part in geom.asMultiPolygon():
                        for point in part.asPolygon()[0]:
                            worksheet.append([point.x(), point.y()])
                else:
                    for point in geom.asPolygon()[0]:
                        worksheet.append([point.x(), point.y()])

            # Save the workbook
            workbook.save(excel_path)

            print(f'Coordinates exported to {excel_path}')
        else:
            print('Operation canceled by user')
    else:
        print('No features selected in the layer')
else:
    print('No active layer in the project')
