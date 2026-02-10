Dim xl, wb
Set xl = CreateObject("Excel.Application")
xl.Visible = False

Set wb = xl.Workbooks.Open("C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main\CONTROLEUR.xlsm")
xl.Run "Batch_MAJ_Exporter_Valeurs"

Set wb = Nothing
Set xl = Nothing