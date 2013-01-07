=begin
	#生成楼幢表
	#城区名称,楼幢表文件，区片商铺数据表文件，基准区片修正体系文件，路段框架表文件
	def Ldb.new_qpldb(cqmz,qpldb_filepath,qpspsj_filepath,jzqpxztx,ldkjbs)

		ldkjb_key=ldkjbs.keys.find{|k|cqmz.include? k}	
		ldkjbsj=nil		
		if ldkjb_key.nil?
			@@logger.error Iconv.conv("gbk", "utf-8","找不到 ")+cqmz+Iconv.conv("gbk", "utf-8"," 的路段框架表")
			return
		else
			ldkjbsj=ldkjbs[ldkjb_key]
		end

		qpspsjb = Qpspsj.load_from_excel(cqmz,qpspsj_filepath)
		
		# 声明Excel的采用的编码为UTF-8
		#WIN32OLE.codepage = WIN32OLE::CP_UTF8

		excel = WIN32OLE::new('excel.Application')	
		excel.Visible = false	
		workbook = excel.Workbooks.Add(1)

		excel.DisplayAlerts = false  

		# 定位于第一个表格
		sheet = workbook.worksheets(1)
		sheet.Select

		rowindex=1
		LDB_COLUMNS.each_with_index do |column_name,index|
			sheet.Cells(rowindex,index+1).value=Iconv.conv("gbk", "utf-8",column_name)
		end

		qpspsjb.each do |qpspsj|
			
			rowindex+=1
			#fix cell values
			FIX_COLUMNS_VALUES.each_pair do | key, value |
				#@@logger.debug key.to_s+","+Iconv.conv("gbk", "utf-8",value)
				sheet.Cells(rowindex,key).value=Iconv.conv("gbk", "utf-8",value)
			end

			#【商铺数据】各条记录的【路段名称】匹配【路段框架表】的【路/项目名称】，
			#如果有找到多个相同名称开头的记录则通过门牌号区别。
			#如果找不到匹配记录则记录下来
			@@logger.debug qpspsj.luduanmc
			if qpspsj.luduanmc.nil?
				@@logger.error Iconv.conv("gbk", "utf-8","此记录没【路段名称】：")+qpspsj.to_s 
			else
				if qpspsj.tdjzmz.nil?
					@@logger.error Iconv.conv("gbk", "utf-8","此记录没【土地/建筑面积】：")+qpspsj.to_s
				else
					ldsj=find_ldkjb(qpspsj,ldkjbsj) 
					if ldsj.nil?
						ldkjbs.each_pair do |key,value|
							next if key.include? ldkjb_key
							ldsj=find_ldkjb(qpspsj,value)
							break unless ldsj.nil?
						end
					end

					@@logger.error Iconv.conv("gbk", "utf-8","无法匹配路段名称：[")+qpspsj.to_s+"]=>"+qpspsj.luduanmc if ldsj.nil?

					unless ldsj.nil?	

						if ldsj.jzjg.nil?
							@@logger.debug Iconv.conv("gbk", "utf-8","此路段基准地价为空值：")+ldsj.lxmmc							
						else		
							jzxx=jzqpxztx.find { |e| e.qpbh.include? ldsj.qpbh }
							if jzxx.nil?
								@@logger.error Iconv.conv("gbk", "utf-8","无法匹配区片编号：")+ldsj.qpbh
							else
								xx=jzxx.get_xx(qpspsj.tdjzmz)/100.to_f

								@@logger.debug Iconv.conv("gbk", "utf-8","修正前的基准价格：")+"#{ldsj.jzjg}"
								ldsj.jzjg=ldsj.jzjg.to_f unless ldsj.jzjg.is_a? Numeric
								jzjg=xx*ldsj.jzjg

								@@logger.debug "#{ldsj.jzjg}*#{xx}=#{jzjg}"

								sheet.Cells(rowindex,20).value=jzjg
							end
						end

						sheet.Cells(rowindex,5).value=ldsj.qpmc
						sheet.Cells(rowindex,6).value=ldsj.qpbh
						sheet.Cells(rowindex,7).value=ldsj.bs
						sheet.Cells(rowindex,8).value=ldsj.lxmmc
						sheet.Cells(rowindex,9).value=ldsj.lxmbh
					end
				end
			end

			sheet.Cells(rowindex,2).value=cqmz
			sheet.Cells(rowindex,10).value=qpspsj.mpdz
			sheet.Cells(rowindex,11).value=qpspsj.mpdz

		end

		@@logger.debug "save as:#{qpldb_filepath}"
		excel.ActiveWorkbook.saveas qpldb_filepath
		excel.ActiveWorkbook.Close(0)
		excel.Quit
	end
	=end