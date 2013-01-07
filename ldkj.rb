#coding: utf-8
require 'win32ole'
require 'logger'
require 'pathname'

#=路段框架
class Ldkj
	#1片区名称,2片区编号,3标识,4路/项目名称,5路/项目编号,6路段起始号,7基准价格
	CQ_EXCEL_INDEX={"白云"=>[4,5,6,7,8,9,12],
		"从化"=>[4,5,6,7,8,9,12],
		"萝岗"=>[4,5,6,7,8,9,12],
		"番禺"=>[4,5,7,6,8,9,12],
		"海珠"=>[4,5,6,7,8,9,12],
		"花都"=>[4,5,6,7,8,9,12],
		"天河"=>[4,5,6,7,8,9,12],
		"荔湾"=>[4,5,6,7,8,9,13],
		"越秀"=>[4,5,6,7,8,9,12],
		"黄埔"=>[4,5,6,7,8,9,12],
		"增城"=>[4,5,6,7,8,9,12],
		"南沙"=>[4,5,6,7,8,9,12]} 
	#行号 片区名称	片区编号	标识	路/项目名称	路/项目编号 基准价格 
	attr_accessor :row,:qpmc,:qpbh,:bs,:lxmmc,:lxmbh,:jzjg,:ldqzh
	def initialize(row,qpmc,qpbh,bs,lxmmc,lxmbh,ldqzh,jzjg)

		#@@logger=Logger.new(STDERR)
		#@@logger.level=Logger::DEBUG

		@row=row
		@qpmc=qpmc
		@qpbh=qpbh
		@bs=bs
		@lxmmc=lxmmc
		@lxmbh=lxmbh
		@jzjg=jzjg
		@ldqzh=ldqzh

		get_ldqzh_map

		@@logger.error Iconv.conv("gbk", "utf-8","基准价格为空值=>[")+to_s+"]" if @jzjg.nil?
	end 

	#从路段起始号生成门牌号范围数组
	def get_ldqzh_map
		return if ldqzh.nil? || ldqzh.empty?
		tmp=@ldqzh.strip.gsub(Iconv.conv("gbk", "utf-8","，"),',')
		@@logger.debug tmp
		ldhs=tmp.split(',')
		@ldmph=Array.new #路段门牌号
		ldhs.each do |ldh|
			hm=ldh.split('-')
			@@logger.debug hm
			@ldmph.push hm[0].to_i..hm[1].to_i
		end
	end

	#是否包含门牌号
	def include_mph(mph)
		unless @ldmph.nil?
			@ldmph.each do |ld|

				return true if ld.include? mph and (ld.first%2)==(mph%2)
			end
		end
		return false
	end

	def to_s
		"#{row},#{qpmc},#{qpbh},#{bs},#{lxmmc},#{lxmbh},#{jzjg},#{ldqzh}"
	end

	#从excel加载区片路段框架数据
	def Ldkj.load_from_excel(cqmz,filepath)		  
		@@logger = Logger.new("log\\#{Pathname.new(filepath).basename}.log")
		@@logger.level=Logger::ERROR
		#@@logger=Logger.new(STDERR)
		#@@logger.level=Logger::DEBUG
		@@logger.formatter = proc { |severity, datetime, progname, msg|
		    "#{msg}\n"
		}

		key=CQ_EXCEL_INDEX.keys.find { |e| cqmz.include? Iconv.conv("gbk", "utf-8",e)  }
		@@logger.error Iconv.conv("gbk", "utf-8","无法匹配区片规则：")+cqmz if key.nil?

		excel=WIN32OLE.new('excel.Application')
		book=excel.Workbooks.open filepath
		sheet=book.Worksheets(1)

		list=Array.new
		(2..sheet.Rows.Count).each do |row|
			break if sheet.Cells(row,1).value.nil? or sheet.Cells(row,1).value.empty?
			data=Ldkj.new(row,sheet.Cells(row, CQ_EXCEL_INDEX[key][0]).value,
				sheet.Cells(row,CQ_EXCEL_INDEX[key][1]).value,
				sheet.Cells(row,CQ_EXCEL_INDEX[key][2]).value,
				sheet.Cells(row, CQ_EXCEL_INDEX[key][3]).value,
				sheet.Cells(row, CQ_EXCEL_INDEX[key][4]).value,
				sheet.Cells(row, CQ_EXCEL_INDEX[key][5]).value,
				sheet.Cells(row, CQ_EXCEL_INDEX[key][6]).value)
		
			list.push data
			@@logger.info data.to_s

		end
		book.saved=false;
		excel.ActiveWorkbook.Close(0)
		excel.Quit
		list
	end
end