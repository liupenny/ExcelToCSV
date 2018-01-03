# xls2csv

## 项目说明

使用apache的poi 实现了 xls/xlsx 和 csv 格式文件的相互转换。

xls/xlsx转换到csv：一个表格中的多个sheet会生成多个csv文件，命名格式：xls/xlsx文件名_Sheet标号.csv

csv(s)转换为xls/xlsx：一个文件夹内必须全部是csv文件，且命名格式按照：xls/xlsx文件名_Sheet标号.csv

### 使用

在TestMain中进行转换调用和调试。

### 项目说明

xls2csv:    xlsx转csv的时候会忽略空格；csv转换回xls/xlsx时不是按顺序的（e.g.正常来说一个xlsx文件中的sheet可能不是按先后顺序的，比如第一个是sheet2,第二个是sheet1. 但当csv转换回去的时候就按照sheet1 sheet2的顺序了）

xls2csv_improved:    xlsx转换csv的时候换了别的库，解决了空格问题；转换顺序问题通过文件前加标号，整合时再删除解决

### 参考资料

http://www.docjar.com/html/api/org/apache/poi/hssf/eventusermodel/examples/XLS2CSVmra.java.html

