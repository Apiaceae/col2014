#!/usr/bin/env ruby

require 'rubygems'
require 'zip/zip' # rubyzip gem
require 'nokogiri'
# require_relative "read_html_content.rb"

class WordXmlManipulate
  def self.open(path, &block)
    self.new(path, &block)
  end

  def initialize(path, &block)
    @replace = {}
    if block_given?
      @zip = Zip::ZipFile.open(path)
      yield(self)
      @zip.close
    else
      @zip = Zip::ZipFile.open(path)
    end
  end

  def merge(rec)
    puts "DEBUG: entering merge function"
    xml = @zip.read("word/document.xml")
    puts "DEBUG: finished reading document"
    doc = Nokogiri::XML(xml) {|x| x.noent}
    puts "DEBUG: finished parsing document with nokogiri"
    (doc/"//w:p").each do |field|
      text_nodeset = (field/".//w:t").first
      if text_nodeset
        if (rec[text_nodeset.inner_html])
          text_nodeset.inner_html = rec[text_nodeset.inner_html].to_s
        end
      end
    end
    @replace["word/document.xml"] = doc.serialize :save_with => 0
  end

  def save(path)
    Zip::ZipFile.open(path, Zip::ZipFile::CREATE) do |output|
      @zip.each do |entry|
        output.get_output_stream(entry.name) do |o|
          if @replace[entry.name]
            o.write(@replace[entry.name])
          else
            o.write(@zip.read(entry.name))
          end
        end
      end
    end
    @zip.close
  end
end

if __FILE__ == $0
  file = ARGV[0]
  working_file = ARGV[1] || file.sub(/\.docx/, '-merged.docx')
  w = WordXmlManipulate.open(file)
  w.merge('ASPLENIACEAE' => '铁角蕨科 Aspleniaceae Newman', 
  'ATHYRIACEAE' => '蹄盖蕨科 Athyriaceae Alston',
  'BLECHNACEAE' => '乌毛蕨科 Blechnaceae Newman',
  'CIBOTIACEAE' => '金毛狗蕨科 Cibotiaceae Korall', 
  'CYATHEACEAE' => '桫椤科 Cyatheaceae Kaulf.',
  'CYSTOPTERIDACEAE' => '冷蕨科 Cystopteridaceae Schmakov',
  'DAVALLIACEAE' => '骨碎补科 Davalliaceae M. R. Schomb.',
  'DENNSTAEDTIACEAE' => '碗蕨科 Dennstaedtiaceae Lotsy',
  'DIPLAZIOPSIDACEAE' => '肠蕨科 Diplaziopsidaceae X. C. Zhang & Christenh.',
  'DIPTERIDACEAE' => '双扇蕨科 Dipteridaceae Seward & E. Dale',
  'DRYOPTERIDACEAE' => '鳞毛蕨科 Dryopteridaceae Herter',
  'EQUISETACEAE' => '木贼科 Equisetaceae Michx. ex DC',
  'GLEICHENIACEAE' => '里白科 Gleicheniaceae C. PreslC. Presl',
  'HYMENOPHYLLACEAE' => '膜蕨科 Hymenophyllaceae Mart.',
  'HYPODEMATIACEAE' => '肿足蕨科 Hypodematiaceae Ching',
  'ISOETACEAE' => '水韭科 Isoetaceae Reichenb.',
  'LINDSAEACEAE' => '鳞始蕨科 Lindsaeaceae C. Presl ex M. R. Schomb.',
  'LOMARIOPSIDACEAE'  => '藤蕨科 Lomariopsidaceae Alston',
  'LYCOPODIACEAE' => '石松科 Lycopodiaceae P. Beauv. ex Mirb.',
  'LYGODIACEAE' => '海金沙科 Lygodiaceae M. Roem.',
  'MARATTIACEAE' => '合囊蕨科 Marattiaceae Kaulf.',
  'MARSILEACEAE' => '蘋科 Marsileaceae Mirb.',
  'NEPHROLEPIDACEAE' => '肾蕨科 Nephrolepidaceae Pic. Serm.', 
  'OLEANDRACEAE' => '条蕨科 Oleandraceae Ching ex Pic. Serm.',
  'ONOCLEACEAE'  => '球子蕨科 Onocleaceae Pic. Serm.', 
  'OPHIOGLOSSACEAE' => '瓶尔小草科 Ophioglossaceae Martinov', 
  'OSMUNDACEAE' => '紫萁科 Osmundaceae Martinov',
  'PLAGIOGYRIACEAE' => '瘤足蕨科 Plagiogyriaceae Bower', 
  'POLYPODIACEAE' => '水龙骨科 Polypodiaceae J. Presl & C. Presl',
  'PSILOTACEAE' => '松叶蕨科 Psilotaceae J. W. Griff. & Henfr.',
  'PTERIDACEAE' => '凤尾蕨科 Pteridaceae E. D. M. Kirchn.', 
  'RHACHIDOSORACEAE' => '轴果蕨科 Rhachidosoraceae X. C. Zhang', 
  'SALVINIACEAE' => '槐叶苹科 Salviniaceae Martinov', 
  'SCHIZAEACEAE' => '莎草蕨科 Schizaeaceae Kaulf.', 
  'SELAGINELLACEAE' => '卷柏科 Selaginellaceae Willk.',
  'TECTARIACEAE' => '三叉蕨科 Tectariaceae Panigrahi',
  'THELYPTERIDACEAE' => '金星蕨科 Thelypteridaceae Pic. Serm.',
  'WOODSIACEAE' => '岩蕨科 Woodsiaceae Herter'
)  
  w.save(working_file)
  puts"DEBUG: complete"
end
