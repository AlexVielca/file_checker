# Script de comprobación de la copia de seguridad del correo electrónico

## Funcionamiento: 

Este script recorre el directorio *\\\vs-nas-01\correos* (indicado en la función **main()**); 

Coge los nombres de las primeras carpetas como los nombres de empleados (excluyendo la carpeta *'Ex-empleados'*);

Obtiene la cantidad de ficheros **.vbk** (que son los 'full backups' de los ficheros de correo), sus fechas de modificación, tamaño, calcula los días desde la última modificación y los muestra en la consola indicando, también, el último fichero .vbk añadido(modificado) y los días desde la última modificación; 

Avisa si un empleado lleva más de 14 días sin un .vbk modificado (creado) con un mensaje: 'Usuarios a revisar: ['Nombre Empleado']';

Tiene la opción de guardar los datos en excel;

Al aceptar la pregunta *'¿Quieres ver el horario de copias (Y/N)?'* se muestran las horas de las copias de seguridad de cada uno de los empleados agrupados por días de la semana.

## Importante:

1. En principio, la opción de eliminar los ficheros .vbk está comentada para poder revisarlos antes de eliminar. Para activar esta opcion hay que descomentar la línea de '**os.remove**'

```
    for vbk_file in vbk_files:                                                          
        if True:                                                                      
            if vbk_file != min_vbk_file:                                                
                # os.remove(vbk_file)
                count_deleted = count_deleted + 1
                print(f"  --x-- Eliminado antiguo: {vbk_file}\n")
```

2. Para activar las instrucciones de guardar excel con los datos hay que descomentar las líneas al final de la función **find_vbk_files()**:

```
    df = pd.DataFrame(data)
    # excel_file = f"vbk_files_{str(datetime.datetime.now())[-5:]}.xlsx"
    excel_file = f"vbk_files_{str(datetime.datetime.now())}.xlsx"
    # df.to_excel(excel_file, index=False)                                          
    # print(f"Excel guardado {excel_file}")
```
