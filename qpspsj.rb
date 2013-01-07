#coding: utf-8
require 'win32ole'
require 'logger'
require 'pathname'

#=区片商铺数据类
#表示【契税含价格数据12.27】目录下的各个表里的一条记录
class Qpspsj
	#门牌地址，楼幢名称，路段名称,门牌号,土地/建筑面积
	attr_accessor :mpdz,:ldmc,:luduanmc,:mph,:tdjzmz,:row
	
	#找楼幢名称规则：匹配字=>匹配字位置偏移量
	FIND_LDMC_BY={"之"=>4,"号"=>1} #,"栋"=>1
	CQ_EXCEL_INDEX={"白云"=>[2,6,7,9],
		"从化"=>[2,8,10,13],
		"萝岗"=>[1,2,1,11],
		"番禺"=>[1,2,1,17],
		"海珠"=>[1,8,10,13],
		"花都"=>[4,8,10,13],
		"天河"=>[1,2,1,12],
		"荔湾"=>[2,8,10,13],
		"越秀"=>[2,8,10,13],
		"黄埔"=>[2,8,10,13],
		"增城"=>[4,8,10,13],
		"南沙"=>[2,8,10,13]} #片区=>excel表位置规则：sheet，门牌地址，路段名称，土地/建筑面积

	def initialize(row,mpdz,luduanmc,tdjzmz)
		@row=row
		@mpdz=mpdz
		@luduanmc=luduanmc
		@tdjzmz=tdjzmz

		find_ldmc	
		find_mph

		@@logger.error Iconv.conv("gbk", "utf-8","门牌地址为空值[")+to_s+"]" if @mpdz.nil?
		@@logger.error Iconv.conv("gbk", "utf-8","路段名称为空值[")+to_s+"]" if @luduanmc.nil?
		@@logger.error Iconv.conv("gbk", "utf-8","土地/建筑面积为空值[")+to_s+"]" if @tdjzmz.nil?
	end 

	def to_s
		"#{@row},#{@mpdz},#{@luduanmc},#{@tdjzmz},#{@ldmc},#{@mph}"
	end

	#从excel加载区片商铺数据
	def Qpspsj.load_from_excel(cqmz,filepath)				  
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
		sheet=book.Worksheets(CQ_EXCEL_INDEX[key][0])

		list=Array.new
		(2..sheet.Rows.Count).each do |row|
			break if sheet.Cells(row,1).value.nil? or sheet.Cells(row,1).value.empty?
			data=Qpspsj.new(row,sheet.Cells(row,CQ_EXCEL_INDEX[key][1]).value,
				sheet.Cells(row, CQ_EXCEL_INDEX[key][2]).value,
				sheet.Cells(row, CQ_EXCEL_INDEX[key][3]).value)
		
			list.push data
			@@logger.info data.to_s

		end
		book.saved=false;
		excel.ActiveWorkbook.Close(0)
		excel.Quit
		list
	end
private
	#找楼幢名称
	def find_ldmc
		FIND_LDMC_BY.each_pair  do |key, value |
			if mpdz.include? Iconv.conv("gbk", "utf-8",key)
				@ldmc_index=mpdz.index(Iconv.conv("gbk", "utf-8",key))
				@ldmc=mpdz[0..@ldmc_index+value]
				break
			end
		end
		if @ldmc.nil?
			@@logger.error Iconv.conv("gbk", "utf-8","无法匹配楼幢名称：[")+to_s+"]=>"+mpdz
		end			
	end

	#找门牌号
	def find_mph
		unless @ldmc.nil?
			start=(/\d/=~@ldmc)

			@mph=@ldmc[start..@ldmc_index-1] unless start.nil?
			@@logger.debug Iconv.conv("gbk", "utf-8","门牌号：")+@mph unless @mph.nil?

		end
	end
end