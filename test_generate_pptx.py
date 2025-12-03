import re
from pathlib import Path
import pandas as pd

CRMPATH = Path('crm.py')
text = CRMPATH.read_text(encoding='utf-8')

# Extract función calcular_analisis_financiero
start1 = text.find('def calcular_analisis_financiero')
if start1 == -1:
    raise SystemExit('no encontrar calcular_analisis_financiero')
next_def = text.find('\ndef ', start1+1)
calc_code = text[start1: next_def]

# Extract formatear_monto
start2 = text.find('def formatear_monto', next_def)
if start2 == -1:
    raise SystemExit('no encontrar formatear_monto')
next_def2 = text.find('\ndef ', start2+1)
form_code = text[start2: next_def2]

# Extract generar_presentacion_dashboard up to def get_base64_image
start3 = text.find('def generar_presentacion_dashboard')
if start3 == -1:
    raise SystemExit('no encontrar generar_presentacion_dashboard')
end_marker = '\ndef get_base64_image'
end3 = text.find(end_marker, start3)
if end3 == -1:
    # fallback: take until end of file
    gen_code = text[start3:]
else:
    gen_code = text[start3:end3]

# Build a namespace and exec
ns = {}
# Provide imports the functions expect
ns['pd'] = pd
ns['Path'] = Path
ns['__file__'] = str(CRMPATH)

# Exec the helper functions
exec(calc_code, ns)
exec(form_code, ns)
exec(gen_code, ns)

# Prepare a sample dataframe
df = pd.DataFrame([
    {'id':'1','nombre':'Alice','sucursal':'CDMX_CENTRO','asesor':'Juan','fecha_ingreso':'01/10/2025','fecha_dispersion':'','estatus':'APROB. CON PROPUESTA','monto_propuesta':'100000','monto_final':''},
    {'id':'2','nombre':'Bob','sucursal':'GDL','asesor':'Ana','fecha_ingreso':'15/10/2025','fecha_dispersion':'','estatus':'PROPUESTA','monto_propuesta':'50000','monto_final':''},
    {'id':'3','nombre':'Carlos','sucursal':'CDMX_CENTRO','asesor':'Juan','fecha_ingreso':'20/10/2025','fecha_dispersion':'05/11/2025','estatus':'DISPERSADO','monto_propuesta':'80000','monto_final':'80000'},
    {'id':'4','nombre':'Diana','sucursal':'MTY','asesor':'Luis','fecha_ingreso':'25/10/2025','fecha_dispersion':'','estatus':'PEND. DOC. PARA EVALUACION','monto_propuesta':'25000','monto_final':''},
])

# Call the function
pptx_bytes = ns['generar_presentacion_dashboard'](df)

out = Path('test_dashboard_output.pptx')
out.write_bytes(pptx_bytes)
print('WROTE', out.resolve())
