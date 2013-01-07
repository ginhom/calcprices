#coding: utf-8
require "iconv" 
require 'fileutils'
load 'qpspsj.rb'
load 'ldb.rb'
load 'ldkj.rb'
load 'jzqpxztx.rb'

if __FILE__ == $0

	#创建日志目录
	if  File.directory? 'log'
		FileUtils.rm_r 'log'
	end
	Dir::mkdir('log')
	
	#exec('taskkill /f /im Excel.exe ') 

	#qpspsj_dir_path="E:\\ghsoft\\calcprices\\契税含价格数据12.27"
	qpspsj_dir_path="E:\\ghsoft\\calcprices\\契税含价格数据12.27"
	ldkjb_dir_path="E:\\ghsoft\\calcprices\\商铺框架修改版本(20130106)"
	ldb_dir_path="E:\\ghsoft\\calcprices\\楼幢表"
	jzqpxztx_filepath="E:\\ghsoft\\calcprices\\广州基准房价商业区片修正体系（十区二市）20121228.xls"
	#filepath=Iconv.conv("gbk", "utf-8","E:\\ghsoft\\calcprices\\路段框架表\\商铺框架表-白云.xls")
	#Ldkj.load_from_excel filepath
	#Jzqpxztx.load_from_excel(jzqpxztx_filepath,Iconv.conv("gbk", "utf-8","天河"))
	#puts filepath
	#filepath=Iconv.conv("gbk", "utf-8","E:\\ghsoft\\calcprices\\契税含价格数据12.27\\越秀区商铺数据.xlsx")
  	#test = Qpspsj.load_from_excel(Iconv.conv("gbk", "utf-8","越秀"),filepath)
  	#ldkj=Ldkj.new(1,"花城广场片","SYGZTH023","商铺楼幢","华穗路2","SYLDTH018",154000,"65-149,176-392")
  	#puts ldkj.include_mph(68)
  	#puts ldkj.include_mph(168)
  	#puts ldkj.include_mph(268)
  	#puts ldkj.include_mph(145)
  	#test_find_mph
  	Ldb.new_from_qpspsj(qpspsj_dir_path,jzqpxztx_filepath,ldkjb_dir_path,ldb_dir_path)
end