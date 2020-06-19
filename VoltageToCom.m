% 水师提督
% 用来生成自定义电压点对应的串口指令（地址 + 端口 + 电压 + CRC16），串口指令中的地址、端口自行设定
% 自定义电压点通过excel下拉生成多个电压值，存放串口指令的一列单元格格式需要事先设置为数字-文本，避免全数字的串口指令受excel写入规则影响
% CRC16检验码生成通过调用函数CRC16_MODBUS
%
file = 'VolToCom.xlsx';% 新建excel文档VolToCom.xlsx，注意存放串口指令的一列单元格格式设置为数字-文本
[num,txt,raw] = xlsread(file);
Voltage = roundn(100*num(:,1),0);% 电压数组从伏特转换为串口指令对应的十进制，这里是*100
for i = 1:length(Voltage)
    com = ['01060003',dec2hex(Voltage(i),4)];% 生成十六进制数据（地址 + 端口 + 电压）
    message = uint16([hex2dec(com(1:2)),hex2dec(com(3:4)),hex2dec(com(5:6)),...
        hex2dec(com(7:8)),hex2dec(com(9:10)),hex2dec(com(11:12))]);% 将十六进制数据每隔8位转换成十进制
    [CRC16] = CRC16_MODBUS(message);% 生成十进制数据对应的校验码
    code = [com,CRC16(2,:),CRC16(1,:)];% 将十六进制数据和校验码合并为串口指令
    rangeB = ['B',num2str(i)];% 设置excel中存放指令位置
    xlswrite(file,{code},1,rangeB);% 串口指令写入excel
end