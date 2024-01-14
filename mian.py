from Flowchart5 import create_graph_from_excel

def main():  #启动
    file_path = 'E:\git clone\excel-to-flowchart\example.xls'
    sheet_name = 'Sheet1'
    DOE_name = 'DOE_name'
    create_graph_from_excel(file_path, sheet_name,DOE_name)

if __name__ == "__main__":
    main()