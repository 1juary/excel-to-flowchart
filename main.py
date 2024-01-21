from Flowchart5 import create_graph_from_excel

def main():  #启动
    file_name = "example"
    file_path = "E:\git clone\excel-to-flowchart"
    fiel_name_xls = file_name + '.xls'
    file_path_with_xls = file_path + '\\' + fiel_name_xls
    sheet_name = 'Sheet1'
    DOE_name = file_name
    create_graph_from_excel(file_path_with_xls, sheet_name, DOE_name)

if __name__ == "__main__":
    main()