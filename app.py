import os
import win32com.client

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IQY_FILE_PATH = os.path.join(BASE_DIR, "query.iqy")

def ler_iqy_para_log():
    print("Iniciando o Excel em segundo plano...")
    excel = None
    wb = None
    try:
        try:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")
        
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)

        print(f"Executando a query usando: {IQY_FILE_PATH}")
        qt = ws.QueryTables.Add(Connection=f"FINDER;{IQY_FILE_PATH}", Destination=ws.Range("A1"))
        qt.BackgroundQuery = False
        qt.Refresh()

        dados = ws.UsedRange.Value or ()

        print("\n" + "="*50)
        print(f"RESULTADO BRUTO ({len(dados)} linhas encontradas):")
        print("="*50 + "\n")

        for i, linha in enumerate(dados):
            # Formata cada linha separada por pipe (|) para visualização crua
            linha_str = " | ".join([str(c) if c is not None else "" for c in linha])
            print(f"Linha {i+1}: {linha_str}")

        print("\n" + "="*50)
        print("Fim dos dados.")

    except Exception as e:
        print(f"Erro ao processar a query: {e}")

    finally:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    ler_iqy_para_log()