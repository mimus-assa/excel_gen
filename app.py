from openpyxl import load_workbook
from flask import Flask, send_from_directory
import os
app = Flask(__name__)

@app.route("/orden/<order>")
def xlsx_row(order):
    #print("importando bases de datos")
    wb_orden = load_workbook(filename = 'orden.xlsx')
    ws_orden = wb_orden["ORDEN"]
    wb_data = load_workbook(filename = "//mnt/winshare/SERV-DEPENDENCIAS (NEWZ).xlsx")
    ws_data = wb_data["SERV-2021"]
    numeric_filter = filter(str.isdigit, order)
    numeric_string = "".join(numeric_filter)
    #print("creando documento")
    dependencia_cel, folio_cell, fecha_cell, equipo_cell, marca_cell, modelo_cell, serie_cell, unidad_cell, delegacion_cell,reparador_cell, servicio_cell, reporte_cell  = "B"+numeric_string, "C"+numeric_string, "M"+numeric_string, "E"+numeric_string, "F"+numeric_string, "G"+numeric_string, "H"+numeric_string, "I"+numeric_string, "J"+numeric_string, "N"+numeric_string, "O"+numeric_string, "P"+numeric_string
    orden_d, dependencia_d, folio_d, fecha_d, equipo_d, marca_d, modelo_d, serie_d, unidad_d, delegacion_d, reparador_d, servicio_d, reporte_d  = ws_data[order].value, ws_data[dependencia_cel].value, ws_data[folio_cell].value, ws_data[fecha_cell].value, ws_data[equipo_cell].value, ws_data[marca_cell].value, ws_data[modelo_cell].value, ws_data[serie_cell].value, ws_data[unidad_cell].value, ws_data[delegacion_cell].value, ws_data[reparador_cell].value, ws_data[servicio_cell].value, ws_data[reporte_cell].value
    ws_orden["H6"], ws_orden["B6"],  ws_orden["H7"], ws_orden["H44"],  ws_orden["B10"], ws_orden["B11"],  ws_orden["B12"], ws_orden["B13"], ws_orden["B16"], ws_orden["B17"], ws_orden["B44"], ws_orden["H11"], ws_orden["A19"] = orden_d, dependencia_d, folio_d, fecha_d, equipo_d, marca_d, modelo_d, serie_d, unidad_d, delegacion_d, reparador_d, servicio_d, reporte_d
    NAME = "generated/"+order+".xlsx"
    wb_orden.save(filename = NAME)
    uploads = os.path.join(app.root_path, "generated")
    return send_from_directory(directory=uploads, filename=order+".xlsx")

    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port='5000', debug=True)    
