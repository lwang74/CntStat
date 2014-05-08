require 'fileutils'
require 'win32ole'

class ExcelConst
	@const_defined = Hash.new
	def self.const_load(object, const_name_space)
		unless @const_defined[const_name_space] then
			WIN32OLE.const_load(object, const_name_space)
			@const_defined[const_name_space] = true
		end
	end
end

module Excel
	def open_read xls_file
		open_excel(xls_file){|excel, workbook|
			yield excel, workbook
		}
	end
	
	def open_rw tmp_xls, dest_xls
		puts "+Writing into Excel file..."
		STDOUT.flush
		open_excel(tmp_xls){|excel, workbook|
			yield workbook
			begin
				#~ p dest_xls
				workbook.saveas "#{FileUtils.pwd}/#{dest_xls}".gsub(/\//, "\\")
				#~ workbook.close true
			rescue WIN32OLERuntimeError =>e
				puts "File '#{FileUtils.pwd}/#{dest_xls}' is using, please close it first!, Enter:Continue, X:Exit."
				STDOUT.flush
				if $stdin.gets=~/^x$/i
				else
					retry
				end
			rescue StandardError =>e
				p e
			#~ ensure
				#~ excel.Application.quit
			end
			#~ excel.quit
		}
	end
protected
	def open_excel xls_file
		excel = WIN32OLE.new('Excel.Application')
		excel.DisplayAlerts = false

		#~ excel.visible = TRUE
		workbook = excel.Workbooks.open('Filename'=>"#{FileUtils.pwd}/#{xls_file}", 'ReadOnly'=>true)
		begin
			yield excel, workbook
		ensure
			workbook.Close
			excel.Application.quit
		end
	end
end

class CExcel
include Excel
end

class CExcel2<CExcel
	def write_area sht, start_cell, arrays
		rg = sht.range(start_cell)
		ExcelConst.const_load(rg, Range)
		rg.EntireRow.Copy()
		tgt_rg = sht.range(sht.Cells(rg.row+1, rg.column), sht.Cells(rg.row+arrays.size-1, rg.column)).EntireRow
		tgt_rg.EntireRow.PasteSpecial(Range::XlPasteFormats)
		arrays.each{|row|
			rg_row = rg
			row.each{|col|
				if col.class==String
					rg_row.value2 = col
				elsif col.class==Hash
					rg_row.value2 = col[:value]
					sht.Hyperlinks.Add rg_row, col[:hyperlink]
				end
				rg_row = rg_row.offset(0, 1)
			}
			rg=rg.offset(1)
		}
		sht.Cells.EntireColumn.AutoFit
	end
end


if __FILE__==$0
	t1 = Time.new
  
	#~ CExcel.new('Template.xlsx', 'abc.xlsx'){|wb|
		#~ wb.Worksheets('Total_parts').range('B6').value='abcxyz'
	#~ }
	
	#~ CExcel.new.open_read('Template.xlsx'){|wb|
		#~ wb.Worksheets(1).usedrange.value2.each{|row|
			#~ p row
		#~ }
	#~ }

	excel = CExcel2.new
	excel.open_rw('config.xlsx', 'output_xls'){|wb|
		sht = wb.worksheets(1)
		excel.write_area sht, 'A2', [['a', 'b'],['c','d'],['cc','dd']]
	}
	puts "time #{Time.new - t1} is spent."
end


