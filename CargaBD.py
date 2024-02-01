#import mariadb
import pandas as pd
import psycopg2
#from datetime import datetime
#import time

archivo = "Informacion.xlsx"
#conexion = mariadb.connect(user="root",password="brandon_1", host="localhost", database = "Bodega")
#print(conexion)

conexion = psycopg2.connect(host="localhost", user="postgres",password="postgres",  database = "Bodega")
print(conexion)
#conexion.autocommit = True

Items=pd.read_excel(archivo,"Items")
Centros=pd.read_excel(archivo,"Centros")
Categorías=pd.read_excel(archivo,"Categorías")
#Colores=pd.read_excel(archivo,"Colores")
#Tipo=pd.read_excel(archivo,"Tipo")
#Clientes=pd.read_excel(archivo,"Clientes")
#Proveedores=pd.read_excel(archivo,"Proveedores")
#Distribucion=pd.read_excel(archivo,"Distribucion")
#Inventario=pd.read_excel(archivo,"Inventario")
#Ventas=pd.read_excel(archivo,"Ventas")
#Compras=pd.read_excel(archivo,"Compras")
for i in range(len(Categorías)):  
    cursor= conexion.cursor()
    Categoría_Código = Categorías.iloc[i]["Categoría (Código)"]
    Categoría_Centro = str(Categorías.iloc[i]["Categoría (Centro)"])
    cursor.execute(f"""INSERT INTO T_Categorias (Categoría_Código,Categoría_Centro) VALUES ('{Categoría_Código}','{Categoría_Centro}');""")
    conexion.commit()

for i in range(len(Centros)):  
    cursor= conexion.cursor()
    Centro = str(Centros.iloc[i]["Centro"])
    Descripción = str(Centros.iloc[i]["Descripción"])
    Categoría_Código = Centros.iloc[i]["Categoría (Código)"]
    Responsable = str(Centros.iloc[i]["Responsable"])
    Area = str(Centros.iloc[i]["Area"])
    if str(Responsable) == "nan":
        Responsable="NULL"
    if str(Area) == "nan":
        Area = 0
    cursor.execute(f"""INSERT INTO T_Centros (Centro,Descripción,Categoría_Código,Responsable,Area_) VALUES ('{Centro}','{Descripción}','{Categoría_Código}','{Responsable}','{Area}');""")
    conexion.commit()

for i in range(len(Items)):  
    cursor= conexion.cursor()
    Código_Artículo = str(Items.iloc[i]["Código Articulo"])
    Artículo = str(Items.iloc[i]["Artículo"])
    Unidad_Medida = str(Items.iloc[i]["Unidad Medida"])
    cursor.execute(f"""INSERT INTO T_Items (Código_Articulo,Artículo,Unidad_Medida) VALUES ('{Código_Artículo}','{Artículo}','{Unidad_Medida}');""")
    conexion.commit()
'''
    for i in range(len(Items)):  
        cursor= conexion.cursor()
        Código_Artículo = str(Items.iloc[i]["Código Articulo"])
        Artículo = str(Items.iloc[i]["Artículo"])
        Unidad_Medida = str(Items.iloc[i]["Unidad Medida"])
        cursor.execute(f'INSERT INTO T_Items (Código_Articulo,Artículo,Unidad_Medida) VALUES ("{Código_Artículo}","{Artículo}","{Unidad_Medida}");')
        conexion.commit()
   
    for i in range(len(Categorías)):  
        cursor= conexion.cursor()
        Categoría_Código = Categorías.iloc[i]["Categoría (Código)"]
        Categoría_Centro = str(Categorías.iloc[i]["Categoría (Centro)"])
        cursor.execute(f'INSERT INTO T_Categorias (Categoría_Código,Categoría_Centro) VALUES ("{Categoría_Código}","{Categoría_Centro}");')
        conexion.commit()
#try:

#except:
#    print(f'{Centro},{Descripción},{Categoría_Código},{Responsable}')
#    print(f'{type(Centro)},{type(Descripción)},{type(Categoría_Código)},{type(Responsable)}')
#finally:
#    print("No se pudo cargar toda la información.")




    for i in range(len(Marca)):
        cursor= conexion.cursor()
        Marca_ = Marca.iloc[i]["Marca"]
        cursor.execute(f'INSERT INTO Marca (Marca) VALUES ("{Marca_}");')
        conexion.commit()

    for i in range(len(Modelo)):
        cursor= conexion.cursor()
        Modelo_ = Modelo.iloc[i]["Modelo"]
        Id_marca = Modelo.iloc[i]["Id_marca"]
        if str(Id_marca) == "nan":
                Id_marca="NULL"
        cursor.execute(f'INSERT INTO Modelo (Modelo, Id_marca) VALUES ("{Modelo_}","{Id_marca}");')
        conexion.commit()

    for i in range(len(Colores)):
        cursor= conexion.cursor()
        Descripcion = Colores.iloc[i]["Descripcion"]
        cursor.execute(f'INSERT INTO Colores (Descripcion) VALUES ("{Descripcion}");')
        conexion.commit()

    for i in range(len(Tipo)):
        cursor= conexion.cursor()
        Descripcion = Tipo.iloc[i]["Descripcion"]
        cursor.execute(f'INSERT INTO Tipo (Descripcion) VALUES ("{Descripcion}");')
        conexion.commit()
    
    for i in range(len(Clientes)):
        cursor= conexion.cursor()
        Nombre = Clientes.iloc[i]["Nombre"]
        Direccion = Clientes.iloc[i]["Direccion"]
        NIT = str(Clientes.iloc[i]["NIT"])
        Telefono = str(Clientes.iloc[i]["Telefono"])
        Correo = str(Clientes.iloc[i]["Correo"])
        Notificaciones = Clientes.iloc[i]["Notificaciones"]
        cursor.execute(f'INSERT INTO Clientes ( Nombre, Direccion, NIT, Telefono, Correo, Notificaciones) VALUES ("{Nombre}","{Direccion}","{NIT}","{Telefono}","{Correo}","{Notificaciones}");')
        conexion.commit()
    
    for i in range(len(Proveedores)):
        cursor= conexion.cursor()
        Nombre = Proveedores.iloc[i]["Nombre"]
        Direccion = Proveedores.iloc[i]["Direccion"]
        cursor.execute(f'INSERT INTO Proveedores ( Nombre, Direccion) VALUES ("{Nombre}","{Direccion}");')
        conexion.commit()
    
    for i in range(len(Distribucion)):
        cursor= conexion.cursor()
        Id_proveedor = Distribucion.iloc[i]["Id_Proveedor"]
        Id_marca = Distribucion.iloc[i]["Id_marca"]
        cursor.execute(f'INSERT INTO Distribucion ( Id_proveedor, Id_Marca) VALUES ({Id_proveedor},{Id_marca});')
        conexion.commit()
    
    for i in range(len(Inventario)):
        cursor= conexion.cursor()
        Id_producto = Inventario.iloc[i]["Id_producto"]
        Id_tipo = Inventario.iloc[i]["Id_tipo"]
        Id_marca = Inventario.iloc[i]["Id_marca"]
        Id_modelo = Inventario.iloc[i]["Id_modelo"]
        Talla = str(Inventario.iloc[i]["Talla"])
        Id_color = Inventario.iloc[i]["Id_color"]
        Existencia = Inventario.iloc[i]["Existencia"]
        Id_sucursal = Inventario.iloc[i]["Id_sucursal"]
        cursor.execute(f'INSERT INTO Inventario ( Id_producto,Id_tipo,Id_marca,Id_modelo, Talla, Id_color,Existencia, Id_sucursal) VALUES ({Id_producto},{Id_tipo},{Id_marca},{Id_modelo},"{Talla}",{Id_color},{Existencia},{Id_sucursal});')
        conexion.commit()

    for i in range(len(Ventas)):
        cursor= conexion.cursor()
        Fecha = Ventas.iloc[i]["Fecha"]
        Id_factura = Ventas.iloc[i]["Id_factura"]   
        Id_cliente = Ventas.iloc[i]["Id_cliente"]
        Id_marca = Ventas.iloc[i]["Id_marca"]
        Id_modelo = Ventas.iloc[i]["Id_modelo"]
        Talla = str(Ventas.iloc[i]["Talla"])
        Id_color = Ventas.iloc[i]["Id_color"]
        Cantidad = Ventas.iloc[i]["Cantidad"]
        Precio = Ventas.iloc[i]["Precio"]
        Total = Ventas.iloc[i]["Total"]
        Id_sucursal = Ventas.iloc[i]["Id_sucursal"]
        Id_tipo = Ventas.iloc[i]["Id_sucursal"]
        Id_inventario = Ventas.iloc[i]["Id_inventario"]
        if str(Id_inventario) == "nan":
            Id_inventario="NULL"
        cursor.execute(f'INSERT INTO Ventas ( Fecha, Id_factura, Id_cliente,Id_marca,Id_modelo, Talla, Id_color,Cantidad, Precio, Total, Id_sucursal, Id_tipo, Id_inventario) VALUES ("{Fecha}", {Id_factura}, {Id_cliente},{Id_marca},{Id_modelo}, "{Talla}", {Id_color},{Cantidad}, {Precio}, {Total}, {Id_sucursal}, {Id_tipo}, {Id_inventario});')
        conexion.commit()

        for i in range(len(Compras)):
            cursor= conexion.cursor()
            Fecha = Compras.iloc[i]["Fecha"]
            Oc_Compra = Compras.iloc[i]["Orden_Compra"]   
            Id_Proveedor = Compras.iloc[i]["Id_Proveedor"]
            Id_marca = Compras.iloc[i]["Id_marca"]
            Id_modelo = Compras.iloc[i]["Id_modelo"]
            Talla = str(Compras.iloc[i]["Talla"])
            Id_color = Compras.iloc[i]["Id_color"]
            Cantidad = Compras.iloc[i]["Cantidad"]
            Precio = Compras.iloc[i]["Precio"]
            Total = Compras.iloc[i]["Total"]
            Id_sucursal = Compras.iloc[i]["Id_sucursal"]
            Id_tipo = Compras.iloc[i]["Id_sucursal"]
            Id_inventario = Compras.iloc[i]["Id_inventario"]
            if str(Id_inventario) == "nan":
                Id_inventario="NULL"
            cursor.execute(f'INSERT INTO Compras ( Fecha, Orden_Compra, Id_proveedor,Id_marca,Id_modelo, Talla, Id_color,Cantidad, Precio, Total, Id_sucursal, Id_tipo, Id_inventario) VALUES ("{Fecha}", {Oc_Compra}, {Id_Proveedor},{Id_marca},{Id_modelo}, "{Talla}", {Id_color},{Cantidad}, {Precio}, {Total}, {Id_sucursal}, {Id_tipo}, {Id_inventario});')
            conexion.commit()


            
'''