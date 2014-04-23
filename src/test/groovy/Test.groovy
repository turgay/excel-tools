import exceltools.CompareExcel

def file1 = "/Developer/projects/excel-tools/src/test/resources/A.xls"
def file2 = "/Developer/projects/excel-tools/src/test/resources/B.xls"

new CompareExcel().compare(file1, file2)
//new CompareExcel().compare(file2, file1)
