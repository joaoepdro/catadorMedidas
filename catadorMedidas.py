import os
from PIL import Image
Image.MAX_IMAGE_PIXELS = None
import openpyxl
from tkinter import filedialog, Tk

root = Tk() 
root.withdraw()
pasta_raiz = filedialog.askdirectory(title="selecione a pasta raiz") # select folder

wb = openpyxl.Workbook() # creation of excel archive
ws = wb.active
ws.append(["Nome do Arquivo", "Caminho Completo", "Largura (cm)", "Altura (cm)", "DPI X", "DPI Y"])

for subdir, dirs, files in os.walk(pasta_raiz): # here the prgram do loop for enters folders enters até nao achar mais
    for file in files:
        if file.lower().endswith('.tif') or file.lower().endswith('.tiff'):
            caminho_completo = os.path.join(subdir, file)
            try:
                with Image.open(caminho_completo) as img:
                    dpi = img.info.get('dpi', (300, 300))
                    dpi_x = dpi[0] if dpi[0] and dpi[0] != 0 else 300
                    dpi_y = dpi[1] if dpi[1] and dpi[1] != 0 else 300
                    width_px, height_px = img.size
                    width_cm = (width_px / dpi_x) * 2.54
                    height_cm = (height_px / dpi_y) * 2.54

                    try:
                        ws.append([
                            str(file),
                            str(caminho_completo),
                            float(round(width_cm, 2)),
                            float(round(height_cm, 2)),
                            float(dpi_x),
                            float(dpi_y)
                        ])
                    except Exception as e:
                        print(f"erro ao salvar {caminho_completo} no excel: {e}")
            except Exception as e:
                print(f"erro ao processar {caminho_completo}: {e}")

# save the excel archive
output_path = os.path.join(pasta_raiz, "tamanhos_tiffs.xlsx")
wb.save(output_path)
print(f"arquivo gerado com sucesso na pasta escolhida {output_path}")
