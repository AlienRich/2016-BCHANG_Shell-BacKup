d:
cd D:\其他資料夾\資訊服務中心-Desktop

rem 總公司
copy twlife.jpg D:\Public

rem 桃園區部
net use u: \\ty0w23001\public
copy twlife.jpg u:\ /y
net use u: /delete

rem 台中第一區部
net use u: \\tc1w23001\public
copy twlife.jpg u:\ /y
net use u: /delete

rem 屏東區部
net use u: \\pt0w23001\public
copy twlife.jpg u:\ /y
net use u: /delete