import os
import datetime
import pandas as pd
import re

def find_vbk_files(directory):
    data = []
    rev_list = []
    count_deleted = 0
    
    day_list = {
        "Monday": [],
        "Tuesday": [],
        "Wednesday": [],
        "Thursday": [],
        "Friday": [],
        "Saturday": [],
        "Sunday": []
    }
    
    for folder in os.listdir(directory):
        folder_path = os.path.join(directory, folder)
        
        if os.path.isdir(folder_path) and not re.search(r"Ex-empleados", folder):                       #   descomentar
            vbk_files = []                                                                            # <-------------+
            for root, dirs, files in os.walk(folder_path):                                              #    comentar   | 
                # vbk_files = []                                                                          # <-------------+--- * Para obtener un total de ficheros .vbk por persona.
                vbk_files.extend([os.path.join(root, file) for file in files if file.endswith('.vbk')]) #               |    (en el caso contrario se busca los .vbk por cada carpeta 
                                                                                                        #   descomentar |        diferente dentro de cada usuario, por ejemplo 
            if vbk_files:                                                                             # <-------------+       si un empleado tiene carpetas Job1 y Job2 para las copias
                if vbk_files:                                                                           #               |       se busca por separado archivos "nuevos" y "viejos" dentro
                    print(f"Nombre: {folder}")                                                          #               |           de cada una de estas carpeta)
                    print(f"Numero de ficheros .vbk: {len(vbk_files)}")                                 #               |
                    print(f"Carpeta: {root}")                                                           #               |
                                                                                                        #               |
                    num_vbk = len(vbk_files)                                                            #               |
                                                                                                        #               |
                    min_days = float('inf')                                                             #               |
                    min_vbk_file = None                                                                 #               |
                                                                                                        #               |
                    for vbk_file in vbk_files:                                                          #               |
                        if True:                                                                      # <- descomentar+
                        # if os.path.dirname(vbk_file) == root:                                           # <- comentar---+
                            creation_time = datetime.datetime.fromtimestamp(os.path.getmtime(vbk_file)) #               |
                            w_day = creation_time.strftime("%A")                                        #               |
                            hour = creation_time.strftime("%H:%M")                                      #               |
                            today = datetime.datetime.today()                                           #               |               
                            time_difference = today - creation_time                                     #               |
                            vbk_file_size = os.path.getsize(vbk_file)                                   #               |
                            size_in_mb = vbk_file_size / 1024 / 1024 # Convertir a Mb                   #               |
                                                                                                        #               |
                            if time_difference.days < min_days:                                         #               |
                                min_days = time_difference.days                                         #               |
                                min_vbk_file = vbk_file                                                 #               |
                                min_creation_date = creation_time                                       #               |
                                min_size_in_mb = size_in_mb                                             #               |
                                                                                                        #               |
                            print(f"  File: {vbk_file}")                                                #               |
                            print(f"  Creation Date: {creation_time}")                                  #               |
                            print(f"  Days Since Creation: {time_difference.days} days. Day: {w_day}")  #               |
                            print(f"  File Size: {size_in_mb:.2f} Mb\n")                                #               |
                                                                                                        #               |
                            # data.append({                                                             #               |
                            #     "Nombre": folder,                                                     #               |
                            #     "Num .vbk": num_vbk,                                                  #               |
                            #     "File Path": vbk_file,                                                #               |
                            #     "Creation Date": creation_time,                                       #               |
                            #     "Days Since Creation": time_difference.days,                          #               |
                            #     "File Size (Mb)": size_in_mb                                          #               |
                            # })                                                                        #               |
                                                                                                        #               |
                    data.append({                                                                       #               |
                        "Nombre": folder,                                                               #               |
                        "Num .vbk": num_vbk,                                                            #               |
                        "Carpeta": root,                                                                #               |
                        "File Path": min_vbk_file,                                                      #               |
                        "Creation Date": min_creation_date,                                             #               |
                        "Creation Day": w_day,                                                          #               |
                        "Days Since Creation": min_days,                                                #               |
                        "File Size (Mb)": min_size_in_mb                                                #               |
                    })                                                                                  #               |
                                                                                                        #               |
                    for vbk_file in vbk_files:                                                          #               |
                        if True:                                                                      # <-descomentar-+
                        # if os.path.dirname(vbk_file) == root:                                           # <-comentar----+
                            if vbk_file != min_vbk_file:                                                
                                # os.remove(vbk_file)                               # <--- Para eliminar el fichero
                                count_deleted = count_deleted + 1
                                print(f"  --x-- Eliminado antiguo: {vbk_file}\n")

                    print(f"    ---> More resent file: {min_vbk_file}")
                    print(f"    ---> With {min_days} days since creation\n")
                    print("-"*30 + "\n")
                    
                    if min_days > 14:
                        rev_list.append(folder)
                    
                    day_list[w_day].append({"Name": folder, "Hour": hour})

    print(f"Ficheros eliminados: {count_deleted}\n")
    print(f"Usuarios a revisar: {rev_list}\n")
    
    if input("¿Quieres ver el horario de copias (Y/N)? ").upper() == "Y":
        print(f" \n")

        for day in day_list:
            print(f"{day} : ") 
            for user in sorted(day_list[day], key=lambda x: x["Hour"]):
                print(f'    {user["Hour"]} - {user["Name"]}, ')
            print(" \n")
        
                   
    df = pd.DataFrame(data)
    # excel_file = f"vbk_files_{str(datetime.datetime.now())[-5:]}.xlsx"
    excel_file = f"vbk_files_{str(datetime.datetime.now())}.xlsx"
    # df.to_excel(excel_file, index=False)                                          # <--- Para guardar en excel
    # print(f"Excel guardado {excel_file}")

def main():
    directory_to_search = r"\\vs-nas-01\correos"  # Cambiar esta línea con la ruta correcta
    find_vbk_files(directory_to_search)

if __name__ == '__main__':
    main()
    input("Presiona Enter para salir...")