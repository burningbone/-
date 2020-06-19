% ˮʦ�ᶽ
% ���������Զ����ѹ���Ӧ�Ĵ���ָ���ַ + �˿� + ��ѹ + CRC16��������ָ���еĵ�ַ���˿������趨
% �Զ����ѹ��ͨ��excel�������ɶ����ѹֵ����Ŵ���ָ���һ�е�Ԫ���ʽ��Ҫ��������Ϊ����-�ı�������ȫ���ֵĴ���ָ����excelд�����Ӱ��
% CRC16����������ͨ�����ú���CRC16_MODBUS
%
file = 'VolToCom.xlsx';% �½�excel�ĵ�VolToCom.xlsx��ע���Ŵ���ָ���һ�е�Ԫ���ʽ����Ϊ����-�ı�
[num,txt,raw] = xlsread(file);
Voltage = roundn(100*num(:,1),0);% ��ѹ����ӷ���ת��Ϊ����ָ���Ӧ��ʮ���ƣ�������*100
for i = 1:length(Voltage)
    com = ['01060003',dec2hex(Voltage(i),4)];% ����ʮ���������ݣ���ַ + �˿� + ��ѹ��
    message = uint16([hex2dec(com(1:2)),hex2dec(com(3:4)),hex2dec(com(5:6)),...
        hex2dec(com(7:8)),hex2dec(com(9:10)),hex2dec(com(11:12))]);% ��ʮ����������ÿ��8λת����ʮ����
    [CRC16] = CRC16_MODBUS(message);% ����ʮ�������ݶ�Ӧ��У����
    code = [com,CRC16(2,:),CRC16(1,:)];% ��ʮ���������ݺ�У����ϲ�Ϊ����ָ��
    rangeB = ['B',num2str(i)];% ����excel�д��ָ��λ��
    xlswrite(file,{code},1,rangeB);% ����ָ��д��excel
end