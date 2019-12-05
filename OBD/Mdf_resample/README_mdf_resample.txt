本程序用于将多个inca采集的.dat文件中的变量根据要求的采样频率进行统一重采样
NOTES:
1. 主程序为Mdf_resample.exe，弹出选择文件对话框时可以多选，读取文件为inca采集的.dat文件
2. 配置文件为Config_resample.xlsx，必须放在此路径下，采样频率在A2单元格，需要重采样的变量名在B列从B2开始往下进行填写或更改，不可填写在其他位置！
3. 输出文件在excel文件夹中，第一个变量time_resample为时间变量（单位：s）


