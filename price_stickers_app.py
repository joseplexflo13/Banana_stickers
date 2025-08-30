import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime


class PriceStickersApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PRICE STICKERS")
        self.root.geometry("400x300")
        self.avios_files = []
        self.order_detail_file = None


        # Botón hoja_consumos_avios_x_OPs
        self.btn_avios = tk.Button(root, text="hoja_consumos_avios_x_OPs", command=self.load_avios_files)
        self.btn_avios.pack(pady=10)


        # Botón orderDetail_Report
        self.btn_order_detail = tk.Button(root, text="orderDetail_Report", command=self.load_order_detail_file)
        self.btn_order_detail.pack(pady=10)


        # Botón PROCESAR
        self.btn_procesar = tk.Button(root, text="PROCESAR", command=self.procesar, state=tk.DISABLED)
        self.btn_procesar.pack(pady=30)


    def load_avios_files(self):
        files = filedialog.askopenfilenames(title="Selecciona archivos Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if files:
            self.avios_files = list(files)
            # Procesar todos los archivos seleccionados y concatenar los resultados
            df_list = []
            for file in self.avios_files:
                try:
                    df = pd.read_excel(file)
                    # Filtrar por columna 'Familia' == 'STICKERS' (columna H)
                    if 'Familia' in df.columns:
                        df = df[df['Familia'] == 'STICKERS']
                        # Filtrar por columna 'Descripcion' conteniendo 'LPN STK' o 'PRICE STICKER' (columna J)
                        if 'Descripcion' in df.columns:
                            mask = df['Descripcion'].astype(str).str.contains('LPN STK', case=False, na=False) | df['Descripcion'].astype(str).str.contains('PRICE STICKER', case=False, na=False)
                            df = df[mask]
                        else:
                            # Si no existe la columna, descartar todo
                            df = df.iloc[0:0]
                        if not df.empty:
                            df_list.append(df)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo procesar el archivo {file}: {e}")
            if df_list:
                df_final = pd.concat(df_list, ignore_index=True)
                # Crear hojas separadas según el contenido de 'Descripcion'
                lpn_stk_df = df_final[df_final['Descripcion'].astype(str).str.contains('LPN STK/', case=False, na=False)].copy()
                price_sticker_df = df_final[df_final['Descripcion'].astype(str).str.contains('PRICE STICKER /', case=False, na=False)].copy()
                # Eliminar el texto inicial correspondiente en la columna 'Descripcion'
                if not lpn_stk_df.empty and 'Descripcion' in lpn_stk_df.columns:
                    lpn_stk_df['Descripcion'] = lpn_stk_df['Descripcion'].astype(str).str.replace(r'^LPN STK/', '', case=False, regex=True).str.strip()
                    # Eliminar columnas Desc_*
                    lpn_stk_df = lpn_stk_df.loc[:, ~lpn_stk_df.columns.str.startswith('Desc_')]
                if not price_sticker_df.empty and 'Descripcion' in price_sticker_df.columns:
                    price_sticker_df['Descripcion'] = price_sticker_df['Descripcion'].astype(str).str.replace(r'^PRICE STICKER / ?', '', case=False, regex=True).str.strip()
                    # Eliminar columnas Desc_*
                    price_sticker_df = price_sticker_df.loc[:, ~price_sticker_df.columns.str.startswith('Desc_')]
                today = datetime.now().strftime('%Y-%m-%d')
                out_name = f"avios_{today}.xlsx"
                try:
                    with pd.ExcelWriter(out_name) as writer:
                        lpn_stk_df.to_excel(writer, sheet_name='LPN_STK', index=False)
                        price_sticker_df.to_excel(writer, sheet_name='PRICE_STICKER', index=False)
                    messagebox.showinfo("Éxito", f"Archivo generado: {out_name}")
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")
            else:
                messagebox.showinfo("Sin datos", "No se encontraron filas con los filtros especificados en los archivos seleccionados.")
            self.check_ready()


    def load_order_detail_file(self):
        file = filedialog.askopenfilename(title="Selecciona archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.order_detail_file = file
            try:
                df = pd.read_excel(file, nrows=0)
                cols = list(df.columns)
                col_str = '\n'.join(cols)
                messagebox.showinfo("Columnas encontradas", f"Columnas en el archivo:\n{col_str}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las columnas: {e}")
            messagebox.showinfo("Carga exitosa", "Se cargó archivo")
            self.check_ready()


    def check_ready(self):
        if self.avios_files and self.order_detail_file:
            self.btn_procesar.config(state=tk.NORMAL)
        else:
            self.btn_procesar.config(state=tk.DISABLED)


    def procesar(self):
        # Nombres de columnas reales para price_stickers
        price_cols = [
            'Type',
            'Order Number',
            'Style Customer Choice Description',
            'Formatted Style Customer Choice Number',
            'Code',
            'Merchandise Ticket Barcode',
            'Size Description',
            'Currency',
            'Amount',
        ]
        # Nombres de columnas reales para lpn
        lpn_cols = [
            'Type',
            'Order Number',
            'Style Customer Choice Description',
            'Size Description',
            'Universal Sku Number',
            'Sku Lpn Barcode',
            'Sku Number',
            'Serialized Barcode Start',
        ]
        if not self.order_detail_file:
            messagebox.showerror("Error", "No se ha cargado el archivo de orderDetail_Report.")
            return
        try:
            # Leer 'Code' como texto para conservar ceros a la izquierda
            df = pd.read_excel(self.order_detail_file, header=0, dtype={"Code": str})
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")
            return
        # price_stickers
        price_row = df[[col for col in price_cols if col in df.columns]].copy()
        for col in price_cols:
            if col not in price_row.columns:
                price_row[col] = ''
        price_row = price_row[price_cols]
        price_row['codigo_avio'] = ''
        price_row['OP'] = ''
        # Forzar columnas a texto en price_stickers
        for col in ["Code", "Merchandise Ticket Barcode"]:
            if col in price_row.columns:
                price_row[col] = price_row[col].astype(str)
        # lpn
        lpn_row = df[[col for col in lpn_cols if col in df.columns]].copy()
        for col in lpn_cols:
            if col not in lpn_row.columns:
                lpn_row[col] = ''
        lpn_row = lpn_row[lpn_cols]
        # Forzar columnas a texto en lpn
        for col in ["Sku Lpn Barcode", "Sku Number"]:
            if col in lpn_row.columns:
                lpn_row[col] = lpn_row[col].astype(str)
        lpn_row['codigo_avio'] = ''
        lpn_row['OP'] = ''
        # Guardar Excel
        today = datetime.now().strftime('%Y-%m-%d')
        out_name = f"lpn_price_{today}.xlsx"
        with pd.ExcelWriter(out_name) as writer:
            price_row.to_excel(writer, sheet_name='price_stickers', index=False)
            lpn_row.to_excel(writer, sheet_name='lpn', index=False)
        messagebox.showinfo("Éxito", f"Archivo generado: {out_name}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PriceStickersApp(root)
    root.mainloop()
