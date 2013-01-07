#coding: UTF-8
require 'win32ole'
require 'logger'
require 'pathname'

class Jzqpxztx
	#行号 片区编号
	attr_accessor :row,:qpbh,:leq60,:gt60_leq72,:gt72_leq96,:gt96_leq150,:gt150

	def initialize(row,qpbh,leq60,gt60_leq72,gt72_leq96,gt96_leq150,gt150)
		@row=row
		@qpbh=qpbh
		@leq60=leq60
		@gt60_leq72=gt60_leq72
		@gt72_leq96=gt72_leq96
		@gt96_leq150=gt96_leq150
		@gt150=gt150
	end

	#通过面积获取修正系数
	def get_xx(area)
		if area<=60
			return leq60
		elsif area>60 and area<=72
			return gt60_leq72
		elsif area>72 and area<=96
			return gt72_leq96
		elsif area>96 and area<=150
			return gt96_leq150
		elsif area>150
			return gt150
		end
	end

	def to_s
		"#{row},#{qpbh},#{leq60},#{gt60_leq72},#{gt72_leq96},#{gt96_leq150},#{gt150}"
	end

	#从excel加载区片路段框架数据
	def Jzqpxztx.load_from_excel(filepath)
		@@logger = Logger.new("log\\#{Pathname.new(Iconv.conv("gbk", "utf-8",filepath)).basename}.log")
		@@logger.level=Logger::ERROR
		#@@logger=Logger.new(STDERR)
		#@@logger.level=Logger::DEBUG
		@@logger.formatter = proc { |severity, datetime, progname, msg|
		    "#{msg}\n"
		}

		excel=WIN32OLE.new('excel.Application')
		book=excel.Workbooks.open(Iconv.conv("gbk", "utf-8",filepath))
		sheetcount=book.Worksheets.count
		list=Array.new
		(1..sheetcount).each do |sheetindex|
			sheetname=book.Worksheets(sheetindex).Name
			@@logger.debug sheetname
			sheet=book.Worksheets(sheetindex) 
			@@logger.debug sheet.Rows.Count
			(4..sheet.Rows.Count).each do |row|
				break if sheet.Cells(row,1).value.nil? or sheet.Cells(row,1).value.empty?
				data=Jzqpxztx.new(row,sheet.Cells(row, 1).value,sheet.Cells(row,37).value,
					sheet.Cells(row, 38).value,sheet.Cells(row, 39).value,
					sheet.Cells(row, 40).value,sheet.Cells(row, 41).value)
			
				list.push data
				@@logger.info Iconv.conv("gbk", "utf-8",data.to_s)

			end
		end

		book.saved=false;
		excel.ActiveWorkbook.Close(0)
		excel.Quit

		list
	end
end