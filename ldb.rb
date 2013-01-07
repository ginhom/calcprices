#coding: UTF-8
require 'logger'
require 'win32ole'

#=楼幢表
class Ldb
	LDB_COLUMNS=["房产类型","城区名称","价格类型","价格形式","片区名称","片区编号",
		"标识","路/项目名称","路/项目编号","楼幢名称","门牌地址","门牌",
		"户号","别名1","别名2","总楼层","临街状况","建筑结构","现状","基准价","评估日期","幢照片"] #栏目
	FIX_COLUMNS_VALUES={1=>"商铺",3=>"楼幢价",4=>"户价格",21=>"2012-7-1"} #栏目=>固定值

	#城区名称,楼幢表文件，区片商铺数据表文件，基准区片修正体系，路段框架表
	def initialize(cqmz,qpldb_filepath,qpspsj_filepath,jzqpxztx,ldkjbs)
		@cqmz=cqmz
		@qpldb_filepath=qpldb_filepath
		@qpspsj_filepath=qpspsj_filepath
		@jzqpxztx=jzqpxztx
		@ldkjbs=ldkjbs

		find_ldkjbsj

		@logger = Logger.new("#{@qpldb_filepath}.log")
		@logger.level=Logger::ERROR
		#@logger=Logger.new(STDERR)
		#@logger.level=Logger::DEBUG
		@logger.formatter = proc { |severity, datetime, progname, msg|
		    "#{msg}\n"
		}
	end

	def make_ldb
		load_qpspsjb

		create_ldb_excel

		add_header

		@qpspsjb.each do |qpspsj|
			
			@rowindex+=1
			
			add_fix_cell_data

			#【商铺数据】各条记录的【路段名称】匹配【路段框架表】的【路/项目名称】，
			#如果有找到多个相同名称开头的记录则通过门牌号区别。
			#如果找不到匹配记录则记录下来
			@logger.debug qpspsj.luduanmc

			if qpspsj.luduanmc.nil?
				@logger.error Iconv.conv("gbk", "utf-8","此记录没【路段名称】：")+qpspsj.to_s 
			else
				if qpspsj.tdjzmz.nil?
					@logger.error Iconv.conv("gbk", "utf-8","此记录没【土地/建筑面积】：")+qpspsj.to_s
				else
					ldsj=try_find_ldsj(qpspsj)

					@logger.error Iconv.conv("gbk", "utf-8","无法匹配路段名称：[")+qpspsj.to_s+"]=>"+qpspsj.luduanmc if ldsj.nil?

					unless ldsj.nil?	

						try_calc_jzjg(ldsj,qpspsj)

						@sheet.Cells(@rowindex,5).value=ldsj.qpmc
						@sheet.Cells(@rowindex,6).value=ldsj.qpbh
						@sheet.Cells(@rowindex,7).value=ldsj.bs
						@sheet.Cells(@rowindex,8).value=ldsj.lxmmc
						@sheet.Cells(@rowindex,9).value=ldsj.lxmbh
					end
				end
			end

			@sheet.Cells(@rowindex,2).value=@cqmz
			@sheet.Cells(@rowindex,10).value=qpspsj.mpdz
			@sheet.Cells(@rowindex,11).value=qpspsj.mpdz

		end

		save_and_close_ldb_excel
	end

	#生成楼幢表
	#区片商铺数据表目录，基准区片修正体系文件，路段框架表目录，楼幢表保存目录
	def Ldb.new_from_qpspsj(qpsp_dir,jzqpxztx_filepath,ldkjb_dir,ldb_dir)
		@@logger = Logger.new("log\\ldb.log")
		@@logger.level=Logger::INFO
		#@@logger=Logger.new(STDERR)
		#@@logger.level=Logger::DEBUG


		Dir::mkdir(Iconv.conv("gbk", "utf-8",ldb_dir)) unless File.directory? Iconv.conv("gbk", "utf-8",ldb_dir)

		qpmz_ldkjb=get_ldkjb(ldkjb_dir) #路段框架表
		qpmz_qpfilepath=get_qpmz(qpsp_dir) #区片商铺数据表

		#加载所有路段框架表
		ldkjbs=Hash.new
		qpmz_ldkjb.each_pair do |key,value|
			ldkjbs[key]=Ldkj.load_from_excel key,value
		end

		#加载所有区片修正体系
		jzqpxztx=Jzqpxztx.load_from_excel(jzqpxztx_filepath)

		#依各区片商铺数据表生成对应区片楼幢表
		qpmz_qpfilepath.each_pair  do |key, value |
			qpldb_filepath=Iconv.conv("gbk", "utf-8",ldb_dir)+"\\"+key+Iconv.conv("gbk", "utf-8","楼幢表.xls")
			@@logger.info qpldb_filepath
			qpldb=Ldb.new(key,qpldb_filepath,value,jzqpxztx,ldkjbs)
			qpldb.make_ldb
		end
	end



	#获取目录下的 区片=>区片数据表名称列表
	def Ldb.get_qpmz(qpsp_dir)
		files= get_dir_files(qpsp_dir)
		qpmz_qpfilepath=Hash.new
		for f in files  
		    if f.include? Iconv.conv("gbk", "utf-8","商铺数据")
		    	qpmz=f[0..f.index(Iconv.conv("gbk", "utf-8","商铺数据"))-1] #区片名字
		    	qpmz_qpfilepath[qpmz]=Iconv.conv("gbk", "utf-8",qpsp_dir)+"\\"+f
		    	@@logger.debug qpmz
		    	@@logger.debug qpmz_qpfilepath[qpmz]
		    end
		end  
		qpmz_qpfilepath
	end 

	#获取目录下的 区片=>路段框架表
	def Ldb.get_ldkjb(ldkjb_dir)
		files= get_dir_files(ldkjb_dir)
		qpmz_ldkjb=Hash.new
		files.each do  |f| 
			path=Pathname.new(f)
			qpmz=f[f.length-path.extname.length-4..f.length-path.extname.length-1]
			qpmz_ldkjb[qpmz]=Iconv.conv("gbk", "utf-8",ldkjb_dir)+"\\"+f
	    	@@logger.debug qpmz
	    	@@logger.debug qpmz_ldkjb[qpmz]
		end
		qpmz_ldkjb
	end

	#获取目录下所有文件
	def Ldb.get_dir_files(dir)
		files=Array.new
		dirp = Dir.open(Iconv.conv("gbk", "utf-8",dir))
		qpmz_qpfilepath=Hash.new
		for f in dirp  
		  case f  
		  when /^\./, /~$/, /\.o/  
		    # do not print  
		  else  
		    @@logger.debug f
		    files.push f
		  end  
		end  
		dirp.close
		files
	end


private


	#计算修正基准价格
	def try_calc_jzjg(ldsj,qpspsj)
		if ldsj.jzjg.nil?
			@logger.debug Iconv.conv("gbk", "utf-8","此路段基准地价为空值：")+ldsj.lxmmc							
		else		
			jzxx=@jzqpxztx.find { |e| e.qpbh.include? ldsj.qpbh }
			if jzxx.nil?
				@logger.error Iconv.conv("gbk", "utf-8","无法匹配区片编号：")+ldsj.qpbh
			else
				xx=jzxx.get_xx(qpspsj.tdjzmz)/100.to_f

				@logger.debug Iconv.conv("gbk", "utf-8","修正前的基准价格：")+"#{ldsj.jzjg}"
				ldsj.jzjg=ldsj.jzjg.to_f unless ldsj.jzjg.is_a? Numeric
				jzjg=xx*ldsj.jzjg

				@logger.debug "#{ldsj.jzjg}*#{xx}=#{jzjg}"

				@sheet.Cells(@rowindex,20).value=jzjg
			end
		end
	end

	#查找当前片区的路段框架表
	def find_ldkjbsj
		@ldkjb_key=@ldkjbs.keys.find{|k|@cqmz.include? k}	
		@ldkjbsj=nil		
		if @ldkjb_key.nil?
			@@logger.error Iconv.conv("gbk", "utf-8","找不到 ")+@cqmz+Iconv.conv("gbk", "utf-8"," 的路段框架表")
		else
			@ldkjbsj=@ldkjbs[@ldkjb_key]
		end
	end

	#加载片区商铺数据
	def load_qpspsjb
		@qpspsjb = Qpspsj.load_from_excel(@cqmz,@qpspsj_filepath)
	end

	#创建楼幢表excel
	def create_ldb_excel
		# 声明Excel的采用的编码为UTF-8
		#WIN32OLE.codepage = WIN32OLE::CP_UTF8

		@excel = WIN32OLE::new('excel.Application')	
		@excel.Visible = false	
		@workbook = @excel.Workbooks.Add(1)

		@excel.DisplayAlerts = false  

		# 定位于第一个表格
		@sheet = @workbook.worksheets(1)
		@sheet.Select
	end

	#保存并关闭楼幢表excel
	def save_and_close_ldb_excel
		@logger.debug "save as:#{@qpldb_filepath}"
		@excel.ActiveWorkbook.saveas @qpldb_filepath
		@excel.ActiveWorkbook.Close(0)
		@excel.Quit
	end

	#添加表头
	def add_header
		@rowindex=1
		LDB_COLUMNS.each_with_index do |column_name,index|
			@sheet.Cells(@rowindex,index+1).value=Iconv.conv("gbk", "utf-8",column_name)
		end
	end

	#添加固定列值
	def add_fix_cell_data
		FIX_COLUMNS_VALUES.each_pair do | key, value |
			#@@logger.debug key.to_s+","+Iconv.conv("gbk", "utf-8",value)
			@sheet.Cells(@rowindex,key).value=Iconv.conv("gbk", "utf-8",value)
		end
	end

	#找出路段数据
	def try_find_ldsj(qpspsj)
		ldsj=find_ldkjb(qpspsj,@ldkjbsj) 
		if ldsj.nil?
			@logger.debug Iconv.conv("gbk", "utf-8","无法在找到对应区片路段数据")
			@ldkjbs.each_pair do |key,value|
				next if key.include? @ldkjb_key
				ldsj=find_ldkjb(qpspsj,value)
				if ldsj.nil?
					@@logger.debug Iconv.conv("gbk", "utf-8","从其它区片路段表找到路段：")+key+"=>"+qpspsj.luduanmc unless ldsj.nil?					
					break
				end
			end
		else
			@logger.debug Iconv.conv("gbk", "utf-8","找到对应区片路段数据")
		end
		ldsj
	end

	#找出路段数据
	def find_ldkjb(qpspsj,ldkjbsj)
		ldsjs=ldkjbsj.find_all do|e| 
			#@logger.debug e.lxmmc+","+qpspsj.luduanmc
			!e.lxmmc.match(/^#{qpspsj.luduanmc.strip}\d?$/).nil?
		end

		ldsj=nil 
		if ldsjs.length==1
			ldsj=ldsjs[0]
		elsif ldsjs.length>1	
			@logger.info Iconv.conv("gbk", "utf-8","找到多个路段:[")+ldsjs.to_s+"]=>"+qpspsj.luduanmc
			mph=qpspsj.mph.to_i
			if mph>0
				ldsj=ldsjs.find { |e| e.include_mph mph  }
			end				
		end	
		return ldsj
	end

end