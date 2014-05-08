require 'yaml'
require './excel'
require 'fileutils'


#~ 学科	年级	课件文件名	作者
#~ 英语	初三	中考英语复习之完型填空及综合填空	吴健
class Total
	def self.init collec
		@total={}
		collec.each{|path, others|
			#~ puts path
			if others[:xls]
				rows = proc_excel("#{path}#{others[:xls]}")
				rows.each{|k, v|
					@total[k]||={}
					v.each{|k1, v1|
						@total[k][k1]||=[]
						v1.each{|one|
							@total[k][k1] << one.update({:path=>path, :ppt=>others[:ppt], :others=>others[:others]})
						}
					}
				}
			end
		}
		#~ p @total
	end
	
	def self.proc_excel excel_file
		rows={}
		CExcel.new.open_read(excel_file){|excel, wb|
			wb.Worksheets.each{|sht|
				#~ puts sht.name
				sht.usedrange.value2.each{|row|
					if row[0]=='学科' and row[1]=='年级' #and row[2]=='课件文件名' and row[3]=='作者'
					elsif row[0]!=''
						rows[row[0]] ||={}
						rows[row[0]][row[1]] ||=[]
						rows[row[0]][row[1]] <<{:name=>row[2], :auth=>row[3]}
					else
					end
				} if sht.usedrange.value2
			}
		}
		rows
	end 
	
	def self.output output_xls, target
		arr = []
		@total.sort.each{|l1|
			#~ puts l1[0]
			FileUtils.mkdir_p "#{target}\\#{l1[0]}", {:mode, :noop}
			l1[1].sort.each{|l2|
				#~ puts "\t#{l2[0]}"
				#~ p l2[1]
				l2[1].sort{|a, b| a[:name]<=>b[:name]}.each{|one|
					#~ puts "\t\t#{one[0]} -> #{one[1]}"
					name = one[:name]
					if one[:ppt]
						target_path = "#{target}\\#{l1[0]}\\#{one[:auth]}"
						FileUtils.mkdir_p target_path, {:mode, :noop}
						target_path_file = "#{target_path}\\#{one[:ppt]}"
						name = {:value=>one[:name], :hyperlink=>target_path_file} 
						FileUtils.cp "#{one[:path]}#{one[:ppt]}", target_path_file
						one[:others].each{|oth|
							p oth
							FileUtils.cp "#{one[:path]}#{oth}", "#{target_path}\\#{oth}"
						} if one[:others]
					end
					arr<<[l1[0], l2[0], name, one[:auth]]
				}
			}
		}

		excel = CExcel2.new
		excel.open_rw('config.xls', output_xls){|wb|
			sht = wb.worksheets(1)
			excel.write_area sht, 'A2', arr
		}
	end
end

def main top_path, target
	collec={}
	Dir["#{top_path}/**/*.*"].each{|one|
		if File.file?(one)
			if one =~ /(.+\/)([^\/]+)$/i
				file_path=$1
				file_name=$2
				collec[file_path]||={}
				if file_name =~ /.+\.xlsx?$/i
					puts one
					STDOUT.flush
					collec[file_path][:xls] = file_name
				elsif file_name =~ /.+\.pptx?$/i
					collec[file_path][:ppt] = file_name
				else
					#~ puts file_name
					collec[file_path][:ppt] = file_name if !collec[file_path][:ppt]
					collec[file_path][:others] ||= []
					collec[file_path][:others]<<file_name
				end
			else
				puts "!Error: #{one}"
			end
		end
	}

	#~ p collec
	Total.init collec

	Total.output '一览表', target
end

main '原稿', '课件'

