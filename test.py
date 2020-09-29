import openpyxl as xl
import table_pack.cleaning as tp

file_name = tp.del_empty_clm('bsa', 4, 'Sheet1')
file_name = tp.choose_clm(file_name, 'Sheet2')
tp.change_exact_column(file_name, 1, 'Sheet3')

