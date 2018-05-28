<?php
/**
 * 常用函数自定义,只支持微擎开发,微擎外请安装yoby开发的Fn扩展类
 * 2018.5.25
 * yoby的微信:logove
 */
/**查找前面字符串中是否包含后者
 * @param $string 原字符串
 * @param $find 子字符串
 * @return bool
 */
 function findStr($string, $find)
{
	return !(strpos($string, $find) === FALSE);
}

/**
 * emoji转换成utf8,方便存储,支持还原
 * de是解码
 */
 function emoji($str, $is = 'en')
{
	if ('en' == $is) {
		if (!is_string($str)) return $str;
		if (!$str || $str == 'undefined') return '';

		$text = json_encode($str);
		$text = preg_replace_callback("/(\\\u[ed][0-9a-f]{3})/i", function ($str) {
			return addslashes($str[0]);
		}, $text);
		return json_decode($text);
	} else {
		$text = json_encode($str);
		$text = preg_replace_callback('/\\\\\\\\/i', function ($str) {
			return '\\';
		}, $text);
		return json_decode($text);
	}
}

/**
 * 时间友好显示,传入时间戳
 */
 function timeline($time)
{
	if (time() <= $time) {
		return date("Y-m-d H:i:s", $time);
	} else {
		$t = time() - $time;
		$f = array(
			'31536000' => '年',
			'2592000' => '个月',
			'604800' => '星期',
			'86400' => '天',
			'3600' => '小时',
			'60' => '分钟',
			'1' => '秒'
		);
		foreach ($f as $k => $v) {
			if (0 != $c = floor($t / (int)$k)) {
				return $c . $v . '前';
			}
		}
	}
}

/**
 * 获取文件扩展名
 * @param $file
 * @return string
 */
 function fileExt($file)
{
	return strtolower(pathinfo($file, 4));
}
/*
获取文件MIME类型,根据文件扩展名来获取
*/
  function getFileType($ext){
	static $mime_types = array (
		'apk'     => 'application/vnd.android.package-archive',
		'3gp'     => 'video/3gpp',
		'ai'      => 'application/postscript',
		'aif'     => 'audio/x-aiff',
		'aifc'    => 'audio/x-aiff',
		'aiff'    => 'audio/x-aiff',
		'asc'     => 'text/plain',
		'atom'    => 'application/atom+xml',
		'au'      => 'audio/basic',
		'avi'     => 'video/x-msvideo',
		'bcpio'   => 'application/x-bcpio',
		'bin'     => 'application/octet-stream',
		'bmp'     => 'image/bmp',
		'cdf'     => 'application/x-netcdf',
		'cgm'     => 'image/cgm',
		'class'   => 'application/octet-stream',
		'cpio'    => 'application/x-cpio',
		'cpt'     => 'application/mac-compactpro',
		'csh'     => 'application/x-csh',
		'css'     => 'text/css',
		'dcr'     => 'application/x-director',
		'dif'     => 'video/x-dv',
		'dir'     => 'application/x-director',
		'djv'     => 'image/vnd.djvu',
		'djvu'    => 'image/vnd.djvu',
		'dll'     => 'application/octet-stream',
		'dmg'     => 'application/octet-stream',
		'dms'     => 'application/octet-stream',
		'doc'     => 'application/msword',
		'dtd'     => 'application/xml-dtd',
		'dv'      => 'video/x-dv',
		'dvi'     => 'application/x-dvi',
		'dxr'     => 'application/x-director',
		'eps'     => 'application/postscript',
		'etx'     => 'text/x-setext',
		'exe'     => 'application/octet-stream',
		'ez'      => 'application/andrew-inset',
		'flv'     => 'video/x-flv',
		'gif'     => 'image/gif',
		'gram'    => 'application/srgs',
		'grxml'   => 'application/srgs+xml',
		'gtar'    => 'application/x-gtar',
		'gz'      => 'application/x-gzip',
		'hdf'     => 'application/x-hdf',
		'hqx'     => 'application/mac-binhex40',
		'htm'     => 'text/html',
		'html'    => 'text/html',
		'ice'     => 'x-conference/x-cooltalk',
		'ico'     => 'image/x-icon',
		'ics'     => 'text/calendar',
		'ief'     => 'image/ief',
		'ifb'     => 'text/calendar',
		'iges'    => 'model/iges',
		'igs'     => 'model/iges',
		'jnlp'    => 'application/x-java-jnlp-file',
		'jp2'     => 'image/jp2',
		'jpe'     => 'image/jpeg',
		'jpeg'    => 'image/jpeg',
		'jpg'     => 'image/jpeg',
		'js'      => 'application/x-javascript',
		'kar'     => 'audio/midi',
		'latex'   => 'application/x-latex',
		'lha'     => 'application/octet-stream',
		'lzh'     => 'application/octet-stream',
		'm3u'     => 'audio/x-mpegurl',
		'm4a'     => 'audio/mp4a-latm',
		'm4p'     => 'audio/mp4a-latm',
		'm4u'     => 'video/vnd.mpegurl',
		'm4v'     => 'video/x-m4v',
		'mac'     => 'image/x-macpaint',
		'man'     => 'application/x-troff-man',
		'mathml'  => 'application/mathml+xml',
		'me'      => 'application/x-troff-me',
		'mesh'    => 'model/mesh',
		'mid'     => 'audio/midi',
		'midi'    => 'audio/midi',
		'mif'     => 'application/vnd.mif',
		'mov'     => 'video/quicktime',
		'movie'   => 'video/x-sgi-movie',
		'mp2'     => 'audio/mpeg',
		'mp3'     => 'audio/mpeg',
		'mp4'     => 'video/mp4',
		'mpe'     => 'video/mpeg',
		'mpeg'    => 'video/mpeg',
		'mpg'     => 'video/mpeg',
		'mpga'    => 'audio/mpeg',
		'ms'      => 'application/x-troff-ms',
		'msh'     => 'model/mesh',
		'mxu'     => 'video/vnd.mpegurl',
		'nc'      => 'application/x-netcdf',
		'oda'     => 'application/oda',
		'ogg'     => 'application/ogg',
		'ogv'     => 'video/ogv',
		'pbm'     => 'image/x-portable-bitmap',
		'pct'     => 'image/pict',
		'pdb'     => 'chemical/x-pdb',
		'pdf'     => 'application/pdf',
		'pgm'     => 'image/x-portable-graymap',
		'pgn'     => 'application/x-chess-pgn',
		'pic'     => 'image/pict',
		'pict'    => 'image/pict',
		'png'     => 'image/png',
		'pnm'     => 'image/x-portable-anymap',
		'pnt'     => 'image/x-macpaint',
		'pntg'    => 'image/x-macpaint',
		'ppm'     => 'image/x-portable-pixmap',
		'ppt'     => 'application/vnd.ms-powerpoint',
		'ps'      => 'application/postscript',
		'qt'      => 'video/quicktime',
		'qti'     => 'image/x-quicktime',
		'qtif'    => 'image/x-quicktime',
		'ra'      => 'audio/x-pn-realaudio',
		'ram'     => 'audio/x-pn-realaudio',
		'ras'     => 'image/x-cmu-raster',
		'rdf'     => 'application/rdf+xml',
		'rgb'     => 'image/x-rgb',
		'rm'      => 'application/vnd.rn-realmedia',
		'roff'    => 'application/x-troff',
		'rtf'     => 'text/rtf',
		'rtx'     => 'text/richtext',
		'sgm'     => 'text/sgml',
		'sgml'    => 'text/sgml',
		'sh'      => 'application/x-sh',
		'shar'    => 'application/x-shar',
		'silo'    => 'model/mesh',
		'sit'     => 'application/x-stuffit',
		'skd'     => 'application/x-koan',
		'skm'     => 'application/x-koan',
		'skp'     => 'application/x-koan',
		'skt'     => 'application/x-koan',
		'smi'     => 'application/smil',
		'smil'    => 'application/smil',
		'snd'     => 'audio/basic',
		'so'      => 'application/octet-stream',
		'spl'     => 'application/x-futuresplash',
		'src'     => 'application/x-wais-source',
		'sv4cpio' => 'application/x-sv4cpio',
		'sv4crc'  => 'application/x-sv4crc',
		'svg'     => 'image/svg+xml',
		'swf'     => 'application/x-shockwave-flash',
		't'       => 'application/x-troff',
		'tar'     => 'application/x-tar',
		'tcl'     => 'application/x-tcl',
		'tex'     => 'application/x-tex',
		'texi'    => 'application/x-texinfo',
		'texinfo' => 'application/x-texinfo',
		'tif'     => 'image/tiff',
		'tiff'    => 'image/tiff',
		'tr'      => 'application/x-troff',
		'tsv'     => 'text/tab-separated-values',
		'txt'     => 'text/plain',
		'ustar'   => 'application/x-ustar',
		'vcd'     => 'application/x-cdlink',
		'vrml'    => 'model/vrml',
		'vxml'    => 'application/voicexml+xml',
		'wav'     => 'audio/x-wav',
		'wbmp'    => 'image/vnd.wap.wbmp',
		'wbxml'   => 'application/vnd.wap.wbxml',
		'webm'    => 'video/webm',
		'wml'     => 'text/vnd.wap.wml',
		'wmlc'    => 'application/vnd.wap.wmlc',
		'wmls'    => 'text/vnd.wap.wmlscript',
		'wmlsc'   => 'application/vnd.wap.wmlscriptc',
		'wmv'     => 'video/x-ms-wmv',
		'wrl'     => 'model/vrml',
		'xbm'     => 'image/x-xbitmap',
		'xht'     => 'application/xhtml+xml',
		'xhtml'   => 'application/xhtml+xml',
		'xls'     => 'application/vnd.ms-excel',
		'xml'     => 'application/xml',
		'xpm'     => 'image/x-xpixmap',
		'xsl'     => 'application/xml',
		'xslt'    => 'application/xslt+xml',
		'xul'     => 'application/vnd.mozilla.xul+xml',
		'xwd'     => 'image/x-xwindowdump',
		'xyz'     => 'chemical/x-xyz',
		'zip'     => 'application/zip'
	);

	return isset($mime_types[$ext]) ? $mime_types[$ext] : 'application/octet-stream';
}
/**
 * 表格转换数组
 * @param $table 表格html代码
 * @return array
 */
 function tableArr($table)
{
	$table = preg_replace("'<table[^>]*?>'si", "", $table);
	$table = preg_replace("'<tr[^>]*?>'si", "", $table);
	$table = preg_replace("'<td[^>]*?>'si", "", $table);
	$table = str_replace("</tr>", "{tr}", $table);
	$table = str_replace("</td>", "{td}", $table);
	//去掉 HTML 标记
	$table = preg_replace("'<[/!]*?[^<>]*?>'si", "", $table);
	//去掉空白字符
	$table = preg_replace("'([rn])[s]+'", "", $table);
	$table = preg_replace('/&nbsp;/', "", $table);
	$table = str_replace(" ", "", $table);
	$table = str_replace(" ", "", $table);

	$table = explode('{tr}', $table);
	array_pop($table);
	foreach ($table as $key => $tr) {
		$td = explode('{td}', $tr);
		array_pop($td);
		$td_array[] = $td;
	}
	return $td_array;
}

/**
 * post提交
 * @param $url
 * @param $msg
 * @return mixed
 */
 function post($url, $msg)
{//post ssl
	$ch = curl_init();
	if (class_exists('\CURLFile')) {
		curl_setopt($ch, CURLOPT_SAFE_UPLOAD, true);
	} else {
		if (defined('CURLOPT_SAFE_UPLOAD')) {
			curl_setopt($ch, CURLOPT_SAFE_UPLOAD, false);
		}
	}
//$msg = array('media'=>"@".$filepath);
//5.6+ $msg = array('media'=>new \CURLFile($filepath));
	preg_match('/https:\/\//', $url) ? $ssl = TRUE : $ssl = FALSE;
	curl_setopt($ch, CURLOPT_POST, 1);
	curl_setopt($ch, CURLOPT_URL, $url);
	if ($ssl) {
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);
		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, FALSE);
	}
	curl_setopt($ch, CURLOPT_POSTFIELDS, $msg);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$data = curl_exec($ch);
	curl_close($ch);
	return $data;
}

/**
 * get 获取
 * @param $url
 * @return mixed
 */
 function get($url)
{
	$ch = curl_init();
	preg_match('/https:\/\//', $url) ? $ssl = TRUE : $ssl = FALSE;
	curl_setopt($ch, CURLOPT_URL, $url);
	curl_setopt($ch, CURLOPT_HEADER, 0);
	if ($ssl) {
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);
		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, FALSE);
	}
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$data = curl_exec($ch);
	curl_close($ch);
	return $data;
}

/**
 * 导出csv数据,不支持大数据,大数据用分页导出
 * $arr = array(
 * array('用户名','密码','邮箱'),
 * array(
 * array('A用户','123456','xiaohai1@zhongsou.com'),
 * array('B用户','213456','xiaohai2@zhongsou.com'),
 * array('C用户','123456','xiaohai3@zhongsou.com')
 * ));
 * putcsv("导出文件",$arr);
 *
 * 导出csv模板
 * $arr = array(array('用户名','密码','邮箱'));
 * putcsv("导出模板",$arr);
 * 文件名不带.csv,自动加
 * $filename 导出文件名
 * $arr 导出数组
 */
 function putcsv($filename, $arr)
{
	if (empty($arr)) {
		return false;
	}
	$export_str = implode(',', $arr[0]) . "\n";

	if (!empty($arr[1])) {
		foreach ($arr[1] as $k => $v) {

			$export_str .= implode(',', $v) . "\n";

		}
	}
	header("Content-type:application/vnd.ms-excel");
	header("Content-Disposition:attachment;filename=" . $filename . date('Y-m-d-H-i-s') . ".csv");
	ob_start();
	ob_end_clean();
	echo "\xEF\xBB\xBF" . $export_str;//解决WPS和excel不乱码
}

/**
 * 导入csv,编码ANSI
 * read.csv数据
 *
 * 商户名称, 昵称, 手机号
 * 惠吃惠喝, 会吃,18291443322
 * egeme, 依加米,18923451622
 * 徐汇区,上海,18291447788
 * 衣服, 买衣服,18291448824
 * 米掌柜, MI,18291448822
 * $path = 'read.csv';
 * $arr= getcsv($path);
 * 导入csv返回数组,注意导入文件一定要是ANSI编码,也就是WPS和excel打开不乱码
 * $path 文件路径
 * */
 function getcsv($path)
{
	$handle = fopen($path, 'r');
	$dataArray = array();
	$row = 0;
	while ($data = fgetcsv($handle)) {
		$num = count($data);

		for ($i = 0; $i < $num; $i++) {
			$dataArray[$row][$i] = mb_convert_encoding($data[$i], "utf-8", 'GBK');
		}
		$row++;

	}

	return $dataArray;
}

/**判断是否微信浏览器
 * @return bool
 */
  function isWeixin() {
	$agent = $_SERVER ['HTTP_USER_AGENT'];
	if (! strpos ( $agent, "icroMessenger" )) {
		return false;
	}
	return true;
}

/**
 * 隐藏手机中间4位
 * @param $phone
 * @return mixed
 */
   function hidetel($phone){
	$IsWhat = preg_match('/(0[0-9]{2,3}[-]?[2-9][0-9]{6,7}[-]?[0-9]?)/i',$phone);
	if($IsWhat == 1){
		return preg_replace('/(0[0-9]{2,3}[-]?[2-9])[0-9]{3,4}([0-9]{3}[-]?[0-9]?)/i','$1****$2',$phone);
	}else{
		return  preg_replace('/(1[34578]{1}[0-9])[0-9]{4}([0-9]{4})/i','$1****$2',$phone);
	}
}
/**
 *
 * @param 生成字符长度 $len
 * @param 生成类型默认大小写数字,0大小写 1数字 2大写 3小写 4中文 $type
 * @param 添加字符后缀 $addChars
 *
 * @return
 */
   function randstr($len=6,$type='',$addChars='') {
	$str ='';
	switch($type) {
		case 0:
			$chars='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'.$addChars;
			break;
		case 1:
			$chars= str_repeat('0123456789',3);
			break;
		case 2:
			$chars='ABCDEFGHIJKLMNOPQRSTUVWXYZ'.$addChars;
			break;
		case 3:
			$chars='abcdefghijklmnopqrstuvwxyz'.$addChars;
			break;
		case 4:
			$chars = "们以我到他会作时要动国产的一是工就年阶义发成部民可出能方进在了不和有大这主中人上为来分生对于学下级地个用同行面说种过命度革而多子后自社加小机也经力线本电高量长党得实家定深法表着水理化争现所二起政三好十战无农使性前等反体合斗路图把结第里正新开论之物从当两些还天资事队批点育重其思与间内去因件日利相由压员气业代全组数果期导平各基或月毛然如应形想制心样干都向变关问比展那它最及外没看治提五解系林者米群头意只明四道马认次文通但条较克又公孔领军流入接席位情运器并飞原油放立题质指建区验活众很教决特此常石强极土少已根共直团统式转别造切九你取西持总料连任志观调七么山程百报更见必真保热委手改管处己将修支识病象几先老光专什六型具示复安带每东增则完风回南广劳轮科北打积车计给节做务被整联步类集号列温装即毫知轴研单色坚据速防史拉世设达尔场织历花受求传口断况采精金界品判参层止边清至万确究书术状厂须离再目海交权且儿青才证低越际八试规斯近注办布门铁需走议县兵固除般引齿千胜细影济白格效置推空配刀叶率述今选养德话查差半敌始片施响收华觉备名红续均药标记难存测士身紧液派准斤角降维板许破述技消底床田势端感往神便贺村构照容非搞亚磨族火段算适讲按值美态黄易彪服早班麦削信排台声该击素张密害侯草何树肥继右属市严径螺检左页抗苏显苦英快称坏移约巴材省黑武培著河帝仅针怎植京助升王眼她抓含苗副杂普谈围食射源例致酸旧却充足短划剂宣环落首尺波承粉践府鱼随考刻靠够满夫失包住促枝局菌杆周护岩师举曲春元超负砂封换太模贫减阳扬江析亩木言球朝医校古呢稻宋听唯输滑站另卫字鼓刚写刘微略范供阿块某功套友限项余倒卷创律雨让骨远帮初皮播优占死毒圈伟季训控激找叫云互跟裂粮粒母练塞钢顶策双留误础吸阻故寸盾晚丝女散焊功株亲院冷彻弹错散商视艺灭版烈零室轻血倍缺厘泵察绝富城冲喷壤简否柱李望盘磁雄似困巩益洲脱投送奴侧润盖挥距触星松送获兴独官混纪依未突架宽冬章湿偏纹吃执阀矿寨责熟稳夺硬价努翻奇甲预职评读背协损棉侵灰虽矛厚罗泥辟告卵箱掌氧恩爱停曾溶营终纲孟钱待尽俄缩沙退陈讨奋械载胞幼哪剥迫旋征槽倒握担仍呀鲜吧卡粗介钻逐弱脚怕盐末阴丰雾冠丙街莱贝辐肠付吉渗瑞惊顿挤秒悬姆烂森糖圣凹陶词迟蚕亿矩康遵牧遭幅园腔订香肉弟屋敏恢忘编印蜂急拿扩伤飞露核缘游振操央伍域甚迅辉异序免纸夜乡久隶缸夹念兰映沟乙吗儒杀汽磷艰晶插埃燃欢铁补咱芽永瓦倾阵碳演威附牙芽永瓦斜灌欧献顺猪洋腐请透司危括脉宜笑若尾束壮暴企菜穗楚汉愈绿拖牛份染既秋遍锻玉夏疗尖殖井费州访吹荣铜沿替滚客召旱悟刺脑措贯藏敢令隙炉壳硫煤迎铸粘探临薄旬善福纵择礼愿伏残雷延烟句纯渐耕跑泽慢栽鲁赤繁境潮横掉锥希池败船假亮谓托伙哲怀割摆贡呈劲财仪沉炼麻罪祖息车穿货销齐鼠抽画饲龙库守筑房歌寒喜哥洗蚀废纳腹乎录镜妇恶脂庄擦险赞钟摇典柄辩竹谷卖乱虚桥奥伯赶垂途额壁网截野遗静谋弄挂课镇妄盛耐援扎虑键归符庆聚绕摩忙舞遇索顾胶羊湖钉仁音迹碎伸灯避泛亡答勇频皇柳哈揭甘诺概宪浓岛袭谁洪谢炮浇斑讯懂灵蛋闭孩释乳巨徒私银伊景坦累匀霉杜乐勒隔弯绩招绍胡呼痛峰零柴簧午跳居尚丁秦稍追梁折耗碱殊岗挖氏刃剧堆赫荷胸衡勤膜篇登驻案刊秧缓凸役剪川雪链渔啦脸户洛孢勃盟买杨宗焦赛旗滤硅炭股坐蒸凝竟陷枪黎救冒暗洞犯筒您宋弧爆谬涂味津臂障褐陆啊健尊豆拔莫抵桑坡缝警挑污冰柬嘴啥饭塑寄赵喊垫丹渡耳刨虎笔稀昆浪萨茶滴浅拥穴覆伦娘吨浸袖珠雌妈紫戏塔锤震岁貌洁剖牢锋疑霸闪埔猛诉刷狠忽灾闹乔唐漏闻沈熔氯荒茎男凡抢像浆旁玻亦忠唱蒙予纷捕锁尤乘乌智淡允叛畜俘摸锈扫毕璃宝芯爷鉴秘净蒋钙肩腾枯抛轨堂拌爸循诱祝励肯酒绳穷塘燥泡袋朗喂铝软渠颗惯贸粪综墙趋彼届墨碍启逆卸航衣孙龄岭骗休借".$addChars;
			break;
		default :
			$chars='ABCDEFGHIJKMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789'.$addChars;
			break;
	}
	if($len>10 ) {
		$chars= $type==1? str_repeat($chars,$len) : str_repeat($chars,5);
	}
	if($type!=4) {
		$chars   =   str_shuffle($chars);
		$str     =   substr($chars,0,$len);
	}else{
		for($i=0;$i<$len;$i++){
			$str.= cutstr($chars,1, floor(mt_rand(0,mb_strlen($chars,'utf-8')-1)),0);
		}
	}
	return $str;
}
/**
 *
 * @param 字符串 $str
 * @param 长度 $length
 * @param 开始位置 $start
 * @param 是否显示... $suffix
 * @param 编码 $charset
 *
 * 截取字符串
 */
if (!function_exists('cutstr')) {
	function cutstr($str, $length, $start = 0, $suffix = true, $charset = "utf-8")
	{
		if (function_exists("mb_substr"))
			$slice = mb_substr($str, $start, $length, $charset);
		elseif (function_exists('iconv_substr')) {
			$slice = iconv_substr($str, $start, $length, $charset);
			if (false === $slice) {
				$slice = '';
			}
		} else {
			$re['utf-8'] = "/[\x01-\x7f]|[\xc2-\xdf][\x80-\xbf]|[\xe0-\xef][\x80-\xbf]{2}|[\xf0-\xff][\x80-\xbf]{3}/";
			$re['gb2312'] = "/[\x01-\x7f]|[\xb0-\xf7][\xa0-\xfe]/";
			$re['gbk'] = "/[\x01-\x7f]|[\x81-\xfe][\x40-\xfe]/";
			$re['big5'] = "/[\x01-\x7f]|[\x81-\xfe]([\x40-\x7e]|\xa1-\xfe])/";
			preg_match_all($re[$charset], $str, $match);
			$slice = join("", array_slice($match[0], $start, $length));
		}
		return $suffix ? $slice . '...' : $slice;
	}
}

/**
 *
 * @param 字节大小 $size
 * @param 保留小数位数 $dec
 *
 * 格式化文件大小
 */
 function getfileSize($size, $dec=2) {
	$a = array("B", "KB", "MB", "GB", "TB", "PB");
	$pos = 0;
	while ($size >= 1024) {
		$size /= 1024;
		$pos++;
	}
	return round($size,$dec)." ".$a[$pos];
}
/**
 *
 * @param 文件名或路径 $file
 *
 * 删除文件夹或文件
 */
   function fileDelete($file){
	if (empty($file))
		return false;
	if (@is_file($file))
		return @unlink($file);
	$ret = true;
	if ($handle = @opendir($file)) {
		while ($filename = @readdir($handle)){
			if ($filename == '.' || $filename == '..')
				continue;
			if (!fileDelete($file . '/' . $filename))
				$ret = false;
		}
	} else {
		$ret = false;
	}
	@closedir($handle);
	if ( file_exists($file) && !rmdir($file) ){
		$ret = false;
	}
	return $ret;
}
/*概率算法
proArr array(100,200,300，400)
function  get_prize(){//获取中奖
$prize_arr = array(
   array('id'=>1,'prize'=>'平板电脑','v'=>1),
   array('id'=>2,'prize'=>'数码相机','v'=>1),
   array('id'=>3,'prize'=>'音箱设备','v'=>1),
  array('id'=>4,'prize'=>'4G优盘','v'=>1),
  array('id'=>5,'prize'=>'10Q币','v'=>1),
  array('id'=>6,'prize'=>'下次没准就能中哦','v'=>95),
);
foreach ($prize_arr as $key => $val) {
   $arr[$val['id']] = $val['v'];
}
$ridk = getrand($arr); //根据概率获取奖项id

$res['yes'] = $prize_arr[$ridk-1]['prize']; //中奖项
unset($prize_arr[$ridk-1]); //将中奖项从数组中剔除，剩下未中奖项
shuffle($prize_arr); //打乱数组顺序
for($i=0;$i<count($prize_arr);$i++){
   $pr[] = $prize_arr[$i]['prize'];
}
$res['no'] = $pr;
return $res;
}
*/
   function getRand($proArr) {
	$result = '';
	$proSum = array_sum($proArr);
	foreach ($proArr as $key => $proCur) {
		$randNum = mt_rand(1, $proSum);
		if ($randNum <= $proCur) {
			$result = $key;
			break;
		} else {
			$proSum -= $proCur;
		}
	}
	unset ($proArr);
	return $result;
}
/**
 * 去除空格 换行
 * @param undefined $str
 *
 * @return
 */
   function trimStr($str)
{
	$str = trim($str);
	$str = preg_replace("/\t/","",$str);
	$str = preg_replace("/\r\n/","",$str);
	$str = preg_replace("/\r/","",$str);
	$str = preg_replace("/\n/","",$str);
	$str = preg_replace("/ /","",$str);
	return trim($str); //返回字符串
}
//去除数组中值两端空格,支持excel
function trimArray($Input){
	if (!is_array($Input))
		return preg_replace("/(\s|\&nbsp\;|　|\xc2\xa0)/","",$Input);
	return array_map('TrimArray', $Input);
}

if (!function_exists('getip')) {
	function getIp()
	{
		$ip = isset($_SERVER['REMOTE_ADDR']) ? $_SERVER['REMOTE_ADDR'] : '';
		if (!preg_match("/^\d+\.\d+\.\d+\.\d+$/", $ip)) {
			$ip = '0';
		}
		return $ip;
	}
}
/**
 * 生成avatar头像
 * @param 邮箱 $email
 * @param 大小 $s
 * @param undefined $d
 * @param undefined $g
 *
 * @return
 */
  function getAvatar($email='', $s=40, $d='mm', $g='g') {
	$hash = md5($email);
	$avatar = "http://www.gravatar.com/avatar/$hash?s=$s&d=$d&r=$g";
	return $avatar;
}

/**
 * 获取内存
 * @return string
 */
  function getMemory(){
	return round((memory_get_usage()/1024/1024),3)."M";
}

/**
 * 加密解密函数
 * ENCODE 加密
 * @param $string
 * @param string $operation
 * @param string $key
 * @param int $expiry
 * @return string
 */
if (!function_exists('authcode')) {
	function authcode($string, $operation = 'DECODE', $key = '', $expiry = 0)
	{
		$ckey_length = 4;
		$keya = md5(substr($key, 0, 16));
		$keyb = md5(substr($key, 16, 16));
		$keyc = $ckey_length ? ($operation == 'DECODE' ? substr($string, 0, $ckey_length) : substr(md5(microtime()), -$ckey_length)) : '';
		$cryptkey = $keya . md5($keya . $keyc);
		$key_length = strlen($cryptkey);
		$string = $operation == 'DECODE' ? base64_decode(substr($string, $ckey_length)) : sprintf('%010d', $expiry ? $expiry + time() : 0) . substr(md5($string . $keyb), 0, 16) . $string;
		$string_length = strlen($string);
		$result = '';
		$box = range(0, 255);
		$rndkey = array();
		for ($i = 0; $i <= 255; $i++) {
			$rndkey[$i] = ord($cryptkey[$i % $key_length]);
		}
		for ($j = $i = 0; $i < 256; $i++) {
			$j = ($j + $box[$i] + $rndkey[$i]) % 256;
			$tmp = $box[$i];
			$box[$i] = $box[$j];
			$box[$j] = $tmp;
		}
		for ($a = $j = $i = 0; $i < $string_length; $i++) {
			$a = ($a + 1) % 256;
			$j = ($j + $box[$a]) % 256;
			$tmp = $box[$a];
			$box[$a] = $box[$j];
			$box[$j] = $tmp;
			$result .= chr(ord($string[$i]) ^ ($box[($box[$a] + $box[$j]) % 256]));
		}
		if ($operation == 'DECODE') {
			if ((substr($result, 0, 10) == 0 || substr($result, 0, 10) - time() > 0) && substr($result, 10, 16) == substr(md5(substr($result, 26) . $keyb), 0, 16)) {
				return substr($result, 26);
			} else {
				return '';
			}
		} else {
			return $keyc . str_replace('=', '', base64_encode($result));
		}
	}
}

/**
 * 生成随机颜色
 * @return string
 */
 function randColor(){
	$char='abcdef0123456789';
	$str='';
	for($i=0;$i<6;$i++){
		$str .= substr($char,mt_rand(0,15),1);
	}
	return '#'.$str;
}
/**
 *
 * @param 数组 $arr
 * @param 层级 $level
 * @param undefined $ptagname
 *
 * 数组转换xml
 */
  function arr2xml($arr, $level = 1, $ptagname = '') {
	$s = $level == 1 ? "<xml>" : '';
	foreach($arr as $tagname => $value) {
		if (is_numeric($tagname)) {
			$tagname = $value['TagName'];
			unset($value['TagName']);
		}
		if(!is_array($value)) {
			$s .= "<{$tagname}>".(!is_numeric($value) ? '<![CDATA[' : '').$value.(!is_numeric($value) ? ']]>' : '')."</{$tagname}>";
		} else {
			$s .= "<{$tagname}>".self::arr2xml($value, $level + 1)."</{$tagname}>";
		}
	}
	$s = preg_replace("/([\x01-\x08\x0b-\x0c\x0e-\x1f])+/", ' ', $s);
	return $level == 1 ? $s."</xml>" : $s;
}
/**
xml转换数组
 */
   function xml2arr($xml) {
	if (empty($xml)) {
		return array();
	}
	$result = array();
	$xmlobj = simplexml_load_string($xml, 'SimpleXMLElement', LIBXML_NOCDATA);
	if($xmlobj instanceof SimpleXMLElement) {
		$result = json_decode(json_encode($xmlobj), true);
		if (is_array($result)) {
			return $result;
		} else {
			return array();
		}
	} else {
		return $result;
	}
}
/*
创建文件或文件夹
参数是数组
["qq/","qq.txt","qqq/tt/"];
*/
  function fileCreate($files) {
	foreach ($files as $key => $value) {
		if(substr($value, -1) == '/'){
			if(!is_dir($value)){mkdir($value, 0777, true);}
		}else{
			if(!file_exists($value)){file_put_contents($value, '');}
		}
	}
}

/**
 *
 * @param undefined $type 弹出类型 1,2,3
 * @param undefined $info 提示语
 * @param undefined $url 跳转地址
 *
 * @return
 */
  function alert($type=1,$info="",$url=""){
	if(1==$type){//自动关闭
		$strs = empty($info)?"":"alert('$info');";
		echo "<script>".$strs."document.addEventListener('WeixinJSBridgeReady', function onBridgeReady() {
WeixinJSBridge.call('closeWindow');});</script>";exit;
	}elseif(2==$type){//显示跳转中...
		$urls = empty($url)?"":'location.href="'.$url.'";';
		$strs = empty($info)?"正在跳转中...":$info;
		echo "<meta charset='utf-8'>";
		die('<script type="text/javascript">document.write("<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,user-scalable=0\"><div style=\"font-size:16px;margin:30px auto;text-align:center;\">'.$strs.'</div>");'.$urls.';</script>');
	}elseif(3==$type){//普通弹出,跳转
		$strs = empty($info)?"":"alert('$info');";
		$urls = empty($url)?"":'location.href="'.$url.'";';
		die('<script type="text/javascript">'.$strs.$urls.'</script>');
	}elseif(4==$type){//蓝色i
		echo "<meta charset='utf-8'>";
		die('<script>document.write("<title>提示</title><meta name=\"viewport\" content=\"width=device-width, initial-scale=1, user-scalable=0\"><link rel=\"stylesheet\"  href=\"https://res.wx.qq.com/open/libs/weui/0.4.3/weui.min.css\"><div class=\"weui_msg\"><div class=\"weui_icon_area\"><i class=\"weui_icon_info weui_icon_msg\"></i></div><div class=\"weui_text_area\"><h4 class=\"weui_msg_title\">'.$info.'</h4></div></div>");document.addEventListener("WeixinJSBridgeReady", function onBridgeReady() {WeixinJSBridge.call("hideOptionMenu");});</script>');

	}elseif(5==$type){//红色警告!
		echo "<meta charset='utf-8'>";
		die('<script>document.write("<title>提示</title><meta name=\"viewport\" content=\"width=device-width, initial-scale=1, user-scalable=0\"><link rel=\"stylesheet\"  href=\"https://res.wx.qq.com/open/libs/weui/0.4.3/weui.min.css\"><div class=\"weui_msg\"><div class=\"weui_icon_area\"><i class=\"weui_icon_msg weui_icon_warn\"></i></div><div class=\"weui_text_area\"><h4 class=\"weui_msg_title\">'.$info.'</h4></div></div>");document.addEventListener("WeixinJSBridgeReady", function onBridgeReady() {WeixinJSBridge.call("hideOptionMenu");});</script>');

	}elseif(6==$type){//绿色成功√
		echo "<meta charset='utf-8'>";
		die('<script>document.write("<title>提示</title><meta name=\"viewport\" content=\"width=device-width, initial-scale=1, user-scalable=0\"><link rel=\"stylesheet\"  href=\"https://res.wx.qq.com/open/libs/weui/0.4.3/weui.min.css\"><div class=\"weui_msg\"><div class=\"weui_icon_area\"><i class=\"weui_icon_msg weui_icon_success\"></i></div><div class=\"weui_text_area\"><h4 class=\"weui_msg_title\">'.$info.'</h4></div></div>");document.addEventListener("WeixinJSBridgeReady", function onBridgeReady() {WeixinJSBridge.call("hideOptionMenu");});</script>');

	}
}
/**
 * 字符串与数组互相转换,转换自动判断是否数组
 */
   function str2arr($var,$str=','){
	if(is_array($var)){
		return implode($str,$var);
	}else{
		return explode($str,$var);
	}
}

/**
 * 判断是否utf8编码
 * @param $string
 * @return bool
 */
  function isutf8($string)
{
	$c    = 0;
	$b    = 0;
	$bits = 0;
	$len  = strlen($string);
	for($i=0; $i<$len; $i++)
	{
		$c = ord($string[$i]);
		if($c > 128)
		{
			if(($c >= 254)) return false;
			elseif($c >= 252) $bits=6;
			elseif($c >= 248) $bits=5;
			elseif($c >= 240) $bits=4;
			elseif($c >= 224) $bits=3;
			elseif($c >= 192) $bits=2;
			else return false;
			if(($i+$bits) > $len) return false;
			while($bits > 1)
			{
				$i++;
				$b=ord($string[$i]);
				if($b < 128 || $b > 191) return false;
				$bits--;
			}
		}
	}
	return true;
}
/**
 * 计算两地之间距离
 * @param undefined $lat1 经度1
 * @param undefined $lng1 纬度1
 * @param undefined $lat2 经度2
 * @param undefined $lng2 纬度2
 * @param 1 $len_type 1是米,2,千米
 * @param undefined $decimal,保留两位小数
 *
 * @return
 */
 function getDistance($lat1, $lng1, $lat2, $lng2, $len_type = 1, $decimal = 2)
{
	$pi = 3.1415926000000001;
	$er = 6378.1369999999997;
	$radLat1 = ($lat1 * $pi) / 180;
	$radLat2 = ($lat2 * $pi) / 180;
	$a = $radLat1 - $radLat2;
	$b = (($lng1 * $pi) / 180) - (($lng2 * $pi) / 180);
	$s = 2 * asin(sqrt(pow(sin($a / 2), 2) + (cos($radLat1) * cos($radLat2) * pow(sin($b / 2), 2))));
	$s = $s * $er;
	$s = round($s * 1000);
	if (1 < $len_type)
	{
		$s /= 1000;
	}
	return round($s, $decimal);
}
/*
* 根据ip获取省市
参数是百度地图key
*/
  function getAddress($ak='') {
	$ak = (empty($ak))?"":$ak;
	$url ="https://api.map.baidu.com/location/ip?ak=$ak&coor=bd09ll&ip=".getip();
	$rs = json_decode(get($url),1);
	if($rs['status']==0){
		$arr = explode('|',$rs['address']);
		$arr = ['province'=>$arr[1],'city'=>$arr[2]];
	}else{
		$arr2 = json_decode(get('http://int.dpool.sina.com.cn/iplookup/iplookup.php?format=json'),1);
		if($arr2['ret']==1){
			$arr = [];
			$arr['country'] = $arr2['country'];
			$arr['province'] = $arr2['province'];
			$arr['city'] = $arr2['city'];
		}
	}
	return $arr;
}

/**
 *
 * @param 代码编号 $code
 * @param 返回的信息 $message
 * @param 数据
 *
 * @return
 */
  function json($code,$message,$list=''){
	$json = [
		'code'=>$code,
		'msg'=>$message,
		'data'=>$list
	];
	header('Content-Type:application/json');
	exit(json_encode($json));


}

/**
 * html转换实体
 */
function htmlencode($var) {
	if (is_array($var)) {
		foreach ($var as $key => $value) {
			$var[htmlspecialchars($key)] = html2($value);
		}
	} else {
		$var = str_replace('&amp;', '&', htmlspecialchars($var, ENT_QUOTES));
	}
	return $var;
}

/**
 * html还原
 */
function htmldecode($var){
	return htmlspecialchars_decode($var);
}

/**
 * 数组编码可存储数据库
 * @param $value
 * @return string
 */
function arrencode($value) {
	return serialize($value);
}

/**
 * 数组解码显示
 * @param $var
 * @return mixed
 */
function arrdecode($var){
	return unserialize($var);
}
/**
 * 生成唯一id字符串
 */
function createid(){
	return md5(uniqid());
}
/**
 *
 * @param 图片路径 $src
 *
 * @return 图片信息
 */
function getimginfo($src){
		$ext = fileext($src);
	if($ext!='gif' && $ext!='png' && $ext!='jpg'){
		return '';
	}
	$srcarr = getimagesize($src);
	$arr = array(1=>'gif','jpg','png');
	return "像素: ".$srcarr[0]." X ".$srcarr[1]." 格式: ".$arr[$srcarr[2]];
}
/**
 * 以下函数是微擎独有的
 */
//分页函数
function pager($tcount, $pindex, $psize = 15, $url = '', $context = array('before' => 5, 'after' => 4, 'ajaxcallback' => '')) {
	global $_W;
	$pdata = array(
		'tcount' => 0,
		'tpage' => 0,
		'cindex' => 0,
		'findex' => 0,
		'pindex' => 0,
		'nindex' => 0,
		'lindex' => 0,
		'options' => ''
	);
	if($context['ajaxcallback']) {
		$context['isajax'] = true;
	}

	$pdata['tcount'] = $tcount;
	$pdata['tpage'] = ceil($tcount / $psize);
	if($pdata['tpage'] <= 1) {
		return '';
	}
	$cindex = $pindex;
	$cindex = min($cindex, $pdata['tpage']);
	$cindex = max($cindex, 1);
	$pdata['cindex'] = $cindex;
	$pdata['findex'] = 1;
	$pdata['pindex'] = $cindex > 1 ? $cindex - 1 : 1;
	$pdata['nindex'] = $cindex < $pdata['tpage'] ? $cindex + 1 : $pdata['tpage'];
	$pdata['lindex'] = $pdata['tpage'];

	if($context['isajax']) {
		if(!$url) {
			$url = $_W['script_name'] . '?' . http_build_query($_GET);
		}
		$pdata['faa'] = 'href="javascript:;" onclick="p(\'' . $_W['script_name'] . $url . '\', \'' . $pdata['findex'] . '\', ' . $context['ajaxcallback'] . ')"';
		$pdata['paa'] = 'href="javascript:;" onclick="p(\'' . $_W['script_name'] . $url . '\', \'' . $pdata['pindex'] . '\', ' . $context['ajaxcallback'] . ')"';
		$pdata['naa'] = 'href="javascript:;" onclick="p(\'' . $_W['script_name'] . $url . '\', \'' . $pdata['nindex'] . '\', ' . $context['ajaxcallback'] . ')"';
		$pdata['laa'] = 'href="javascript:;" onclick="p(\'' . $_W['script_name'] . $url . '\', \'' . $pdata['lindex'] . '\', ' . $context['ajaxcallback'] . ')"';
	} else {
		if($url) {
			$pdata['faa'] = 'href="?' . str_replace('*', $pdata['findex'], $url) . '"';
			$pdata['paa'] = 'href="?' . str_replace('*', $pdata['pindex'], $url) . '"';
			$pdata['naa'] = 'href="?' . str_replace('*', $pdata['nindex'], $url) . '"';
			$pdata['laa'] = 'href="?' . str_replace('*', $pdata['lindex'], $url) . '"';
		} else {
			$_GET['page'] = $pdata['findex'];
			$pdata['faa'] = 'href="' . $_W['script_name'] . '?' . http_build_query($_GET) . '"';
			$_GET['page'] = $pdata['pindex'];
			$pdata['paa'] = 'href="' . $_W['script_name'] . '?' . http_build_query($_GET) . '"';
			$_GET['page'] = $pdata['nindex'];
			$pdata['naa'] = 'href="' . $_W['script_name'] . '?' . http_build_query($_GET) . '"';
			$_GET['page'] = $pdata['lindex'];
			$pdata['laa'] = 'href="' . $_W['script_name'] . '?' . http_build_query($_GET) . '"';
		}
	}

	$html = '	<div class="page-hd bg-gray" style="height:32px;">
	<div class="pager"  id="pager"><div class="pager-left">';

	$html .= "<div class=\"pager-first\"><a {$pdata['faa']} class=\"pager-nav\">首页</a></div>";
	$html .= "<div class=\"pager-pre\"><a {$pdata['paa']} class=\"pager-nav\">上一页</a></div>";
	$html .='</div><div class="pager-cen">
					' .$pindex.'/'.$pdata['tpage'].'
				</div><div class="pager-right">';

	$html .= "<div class=\"pager-next\"><a {$pdata['naa']} class=\"pager-nav\">下一页</a></div>";
	$html .= "<div class=\"pager-end\"><a {$pdata['laa']} class=\"pager-nav\">尾页</a></div>";

	$html .= '</div></div></div>';
	return $html;
}
/*
生成唯一订单号
表名 ,字段名,前缀
*/
function ordersn($table, $field, $prefix='')
{
	$billno = date('YmdHis') . randstr(6,1);

	while (1) {
		$count = pdo_fetchcolumn('select count(*) from ' . tablename($table) . ' where ' . $field . '=:billno limit 1', array(':billno' => $billno));

		if ($count <= 0) {
			break;
		}

		$billno = date('YmdHis') .randstr(6,1);
	}

	return $prefix . $billno;
}
/**
 *
 * @param  $openid
 * 判断用户是否关注,不填写openid自动获取当前用户openid,关注返回1,其他都返回0
 * @return
 */
function isfollow($openid = '')
{
	global $_W;
	$openid = (empty($openid))?$_W['openid']:$openid;
	if (!empty($openid))
	{
		$rs = pdo_fetch('select follow from ' . tablename('mc_mapping_fans') . ' where openid=:openid and uniacid=:uniacid limit 1', array(':openid' => $openid, ':uniacid' => $_W['uniacid']));
		$followed = ($rs['follow'] == 1)?1:0;
		return $followed;
	}else{
		return 0;
	}

}
/**
 * 生成日志
 */
function logging($str){
	load()->func('logging');
	logging_run($str);
}
/**
 * 通过Oauth获取用户信息
 * @param undefined $appid
 * @param undefined $secret
 * @param undefined $snsapi,snsapi_userinfo
 * @param undefined $expired
 *
 * @return
 */
function getoauth($appid, $secret, $snsapi = 'snsapi_base', $expired = '600')
{
	global $_W;
	$wxuser = $_COOKIE[$_W['config']['cookie']['pre'] . $appid];
	if ($wxuser === NULL)
	{
		$code = ((isset($_GET['code']) ? $_GET['code'] : ''));
		if (!($code))
		{
			$url =$_W['siteurl'];
			$oauth_url = 'https://open.weixin.qq.com/connect/oauth2/authorize?appid=' . $appid . '&redirect_uri=' . urlencode($url) . '&response_type=code&scope=' . $snsapi . '&state=wxbase#wechat_redirect';
			header('Location: ' . $oauth_url);
			exit();
		}
		load()->func('communication');
		$getOauthAccessToken = ihttp_get('https://api.weixin.qq.com/sns/oauth2/access_token?appid=' . $appid . '&secret=' . $secret . '&code=' . $code . '&grant_type=authorization_code');
		$json = json_decode($getOauthAccessToken['content'], true);
		if (!(empty($json['errcode'])) && (($json['errcode'] == '40029') || ($json['errcode'] == '40163')))
		{
			$url = $_W['siteurl'];
			$parse = parse_url($url);
			if (isset($parse['query']))
			{
				parse_str($parse['query'], $params);
				unset($params['code']);
				unset($params['state']);
				$url = $_W['siteroot'] . $parse['path'] . '?' . http_build_query($params);
			}
			$oauth_url = 'https://open.weixin.qq.com/connect/oauth2/authorize?appid=' . $appid . '&redirect_uri=' . urlencode($url) . '&response_type=code&scope=' . $snsapi . '&state=wxbase#wechat_redirect';
			header('Location: ' . $oauth_url);
			exit();
		}
		if ($snsapi == 'snsapi_userinfo')
		{
			$userinfo = ihttp_get('https://api.weixin.qq.com/sns/userinfo?access_token=' . $json['access_token'] . '&openid=' . $json['openid'] . '&lang=zh_CN');
			$userinfo = $userinfo['content'];
		}
		else if ($snsapi == 'snsapi_base')
		{
			$userinfo = array();
			$userinfo['openid'] = $json['openid'];
		}
		$userinfostr = json_encode($userinfo);
		isetcookie($appid, $userinfostr, $expired);
		return $userinfo;
	}
	return json_decode($wxuser, true);
}

/*生成二维码文件*/
function qr($url)
{
	global $_W;
	$path = IA_ROOT . '/attachment/images/qrcode/' . $_W['uniacid'] . '/';
	if (!(is_dir($path)))
	{
		load()->func('file');
		mkdirs($path);
	}
	$file = md5(base64_encode($url)) . '.jpg';
	$qrcode_file = $path . $file;
	if (!(is_file($qrcode_file)))
	{
		require_once IA_ROOT . '/framework/library/qrcode/phpqrcode.php';
		QRcode::png($url, $qrcode_file, QR_ECLEVEL_L, 4);
	}
	return $_W['siteroot'] . 'attachment/images/qrcode/' . $_W['uniacid'] . '/' . $file;
}
/*
 * 根据openid,获取用户信息
 */
function getuserinfo($openid){
	global $_W;
	load()->classs('weixin.account');
	$acc = WeiXinAccount::create($_W['acid']);
	return $acc->fansQueryInfo($openid);
}
/*
 * 根据媒体id下载图片
 * 返回图片地址
 */
function downimg($mid){
	global $_W;
	load()->classs('weixin.account');
	$acc = WeiXinAccount::create($_W['acid']);
	$arr = ['media_id'=>$mid,'type'=>"image"];
	return $acc->downloadMedia($arr);
}
/**
 * 上传图片到微信服务器
 * 返回数组
 */
function upimg($filename){
	global $_W;
	load()->classs('weixin.account');
	$acc = WeiXinAccount::create($_W['acid']);
	return $acc->uploadMedia($filename);
}
/**
 * 长网址转换短网址
 */
function shorturl($url){
	global $_W;
	load()->classs('weixin.account');
	$acc = WeiXinAccount::create($_W['acid']);
	return $acc->long2short($url)['short_url'];
}
/**
 * 发送文本客服消息
 */
function sendtxt($openid,$str){
	$message = [
		'msgtype'=>'text',
		'text'=>['content'=>urlencode($str)],
		'touser'=>$openid
	];
	$acc = WeAccount::create();
	$status = $acc->sendCustomNotice($message);
	if (is_error($status)) {
		message('发送失败，原因为' . $status['message']);
	}
}

/**
 * 发送图片客服消息
 */
function sendimg($openid,$mid){
	$message = [
		'touser' => $openid,
		'msgtype' => 'image',
		'image' =>['media_id'=> $mid]
	];
	$acc = WeAccount::create();
	$status = $acc->sendCustomNotice($message);
	if (is_error($status)) {
		message('发送失败，原因为' . $status['message']);
	}
}

/**
 * @param $openid
 * @param $tplid 模板id
 * @param $tplstr 模板内容
 * @param $content 发送内容数组
 * @param $url 跳转地址
 */
function sendtpl($openid,$tplid,$tplstr,$content,$url){

	$arr  =[];
	preg_match_all('/{{(.*).DATA}}/',$tplstr,$rs);
	foreach($rs[1] as  $k=>$v){
		$arr[$v] = ['value'=>$content[$k]];
	}
	$arr['first']['color']='#04be02';
	$arr['remark']['color']='#18b4ed';
	$acc = WeAccount::create();
	$status = $acc->sendTplNotice($openid,$tplid, $arr, $url); //1表示发送成功
	return $status;
}

/**
 * 积分操作
 * @param uid或openid
 * @param string $type credit1积分 2余额
 * @param $val 积分支持正负
 * @param string $str 备注
 * @return bool
 */
function credit($uid, $type='credit1', $n,$str="积分操作备注"){
		global $_W;
	load()->model('mc');
	checkauth();
	if (is_string($uid)) {
		$uid = mc_openid2uid($uid);
	}
	$log = [0,$str];
	mc_credit_update($uid, $type, $n, $log);
return true;
}

/**
 * 查询积分
 * @param $uid
 * @param array $type 类型数组
 * @return mixed
 */
function getcredit($uid,$type=['credit1']){
	global $_W;
	load()->model('mc');
	if (is_string($uid)) {
		$uid = mc_openid2uid($uid);
	}
	return mc_credit_fetch($uid,$type);
}
//echo '<pre>'.print_r(get_defined_functions()['user'],1).'</pre>';
