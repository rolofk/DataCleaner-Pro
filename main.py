import pandas as pd
import os

def limpiar_datos_ventas(archivo_entrada, archivo_salida):
    print("🚀 Iniciando DataCleaner Pro...")
    print("-" * 40)
    
    # Verificamos si el archivo de origen realmente existe
    if not os.path.exists(archivo_entrada):
        print(f"❌ Error: No se encontró el archivo '{archivo_entrada}'.")
        return

    try:
        # 1. Leer los datos crudos
        print("📖 Leyendo archivo desordenado...")
        df = pd.read_csv(archivo_entrada)
        total_inicial = len(df)
        
        # 2. Eliminar filas duplicadas exactamente iguales
        print("🧹 Eliminando registros duplicados...")
        df = df.drop_duplicates()
        
        # 3. Rellenar valores vacíos de forma inteligente
        print("🩹 Corrigiendo valores en blanco...")
        df['Cliente'] = df['Cliente'].fillna("Cliente Desconocido")
        df['Precio'] = df['Precio'].fillna(0.0)
        
        # 4. Estandarizar el formato de texto (mayúsculas y minúsculas)
        print("📝 Estandarizando formato de texto...")
        df['Cliente'] = df['Cliente'].str.title() # Convierte "juan perez" a "Juan Perez"
        df['Producto'] = df['Producto'].str.capitalize()
        
        # 5. Exportar el resultado a un Excel limpio
        print("💾 Generando reporte final en Excel...")
        df.to_excel(archivo_salida, index=False, engine='openpyxl')
        
        total_final = len(df)
        print("-" * 40)
        print("✅ ¡Éxito! Proceso terminado satisfactoriamente.")
        print(f"📊 Resumen: De {total_inicial} filas originales, quedaron {total_final} filas limpias.")
        print(f"📁 Archivo guardado como: {archivo_salida}")
        
    except Exception as e:
        print(f"❌ Ocurrió un error inesperado durante el proceso: {e}")

# Comando de ejecución principal
if __name__ == "__main__":
    limpiar_datos_ventas("ventas_crudas.csv", "reporte_limpio.xlsx")