import xlrd
import graphviz
def create_graph_from_excel(file_path, sheet_name, DOE_name):
    graph = graphviz.Digraph()
    graph.graph_attr['label'] = DOE_name
    graph.graph_attr['labelloc'] = 't'
    graph.graph_attr['fontname'] = 'SimSun'  # 设置字体为宋体或其他支持中文的字体
    # ExcelFileName = 'example.xls'
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(sheet_name)

    num_rows = worksheet.nrows
    num_cols = worksheet.ncols

    result_data = []
    for curr_row in range(0, num_rows, 1):
        row_data = []

        for curr_col in range(0, num_cols, 1):
            data = worksheet.cell_value(curr_row, curr_col)  # Read the data in the current cell
            # print(data)
            data = data.encode('utf-8').decode('utf-8')
            row_data.append(data)

        result_data.append(row_data)


    node = result_data
    noderesult = []*(num_rows*num_cols)
    for i in range(len(result_data)):
        for j in range(len(result_data[0])):
            node[i][j] = str(result_data[i][j])
            if((node[i][j])!=""):
                graph.node(node[i][j], style="rounded, filled", shape="box", fontname="SimSun") #在有向图中增加节点，node 是list格式
                noderesult.append(node[i][j])

    for t in range( len(node[0])-1):
        if (node[0][t]!=''):
            graph.edge(node[0][t], node[0][t+1], dir="forward", arrowhead="normal", style="")



    for p in range(len(result_data[0])): #列
        for q in range(1, len(result_data)): #行
            if(node[q][p]!=''):
                graph.edge(node[q][p], node[0][p+1], dir="forward", arrowhead="normal", style="")


    graph.render('example1_graph', format='png')