# we7fn
微擎常用函数分享,包括常用类的使用

# 用法
~~~
excel的使用,支持导入xls.xlsx,csv导出支持xls,xlsx

	$csv = new Excel();
//导入同时支持三种格式,需要判断文件格式,返回数组,去除第一行表头.
	if($ext!='csv'){
					$arr = $csv->loadExcel($new)->getSheet(0)->toArray();
				}else{
					$arr = $csv->loadExcel($new);
				}
array_shift($arr);

//创建导出数据,第一是个参数是表名称,第二个是导出数据,为空的时候是导出模板,必须是数组格式,第三个是excel表头,必须是数组,
$csv->createSheet('导出模板',[],$arr);

//弹出下载,第一个文件名,第二个xls或xlsx
$csv->createSheet('导出模板',[],$arr)->downFile($filename,'xls');


//导出到服务器,参数分别是文件名,保存路径,为空时候服务器data/Excel/下,第三个xls,或xlsx返回不同excel.
$csv->createSheet('导出模板',[],$arr)->saveFile($excelName='',$filepath='',$ver='xls')
~~~

~~~
global.php公共函数库,有些和微擎是重复的,但是判断了,所以不会存在冲突定义,无需删除,

Array
(
    [0] => findstr 查找子字符串
    [1] => emoji emoji转换与还原
    [2] => timeline 友好时间显示
    [3] => fileext 获取文件扩展名
    [4] => getfiletype 根据扩展名获取MIME
    [5] => tablearr 表格转数组
    [6] => post POSTcurl
    [7] => get 
    [8] => putcsv 输出csv
    [9] => getcsv 导入csv
    [10] => isweixin 是否微信浏览器
    [11] => hidetel 隐藏手机号中间四位
    [12] => randstr 随即字符串
    [13] => getfilesize 字节转换GB MB
    [14] => filedelete 删除文件
    [15] => getrand 抽奖算法
    [16] => trimstr 去除字符串空格
    [17] => trimarray 去除数组空格支持excel
    [18] => getavatar 获取头像
    [19] => getmemory 获取内存
    [20] => randcolor 随机颜色值
    [21] => arr2xml 数组转xml
    [22] => xml2arr xml转数组
    [23] => filecreate 创建文件或文件夹
    [24] => alert 提示信息
    [25] => str2arr 字符数组互相转换
    [26] => isutf8 是否utf8编码
    [27] => getdistance 计算经纬度距离
    [28] => getaddress 根据ip获取省市县
    [29] => json 输出json格式
    [30] => htmlencode html转实体
    [31] => htmldecode 实体还原
    [32] => arrencode 数组序列化
    [33] => arrdecode 反序列化
    [34] => createid 创建随机文件名
    [35] => getimginfo 获取图片像素
    [36] => pager 分页
    [37] => ordersn 唯一id生成,订单号生成
    [38] => isfollow 是否关注
    [39] => logging 日志生成
    [40] => getoauth OAUth2获取信息
    [41] => qr 二维码生成
    [42] => getuserinfo 根据openid
    [43] => downimg 从微信下载图片
    [44] => upimg 上传到微信服务器图片
    [45] => shorturl 长网址转换短网址
    [46] => sendtxt 发送文本客服消息
    [47] => sendimg 发送图片客服消息
    [48] => sendtpl 发送模板消息
    [49] => credit 积分操作
    [50] => getcredit 查询积分
    [51] => cutstr 截取字符串
    [52] => getip 获取ip
    [53] => authcode 加密解密字符串
)

~~~
