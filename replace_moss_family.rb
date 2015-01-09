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
  w.merge('ACROBOLBACEAE' => '顶苞苔科 Acrobolbaceae E. A. Hodgs.', 'ADELANTHACEAE' =>  '隐蒴苔科 Adelanthaceae Grolle', 
  'ALLISONIACEAE' => '苞叶苔科 Allisoniaceae (R. M. Schust. ex Grolle) Schljakov', 
  'AMBLYSTEGIACEAE' => '柳叶藓科 Amblystegiaceae G. Roth',  
  'AMPHIDIACEAE' => '瓶藓科 Amphidiaceae M. Stech', 
  'ANASTROPHYLLACEAE' => '挺叶苔科 Anastrophyllaceae S?derstr & Roo & Hedd.',
'ANDREAEACEAE' => '黑藓科 Andreaeaceae Dumort.',
'ANEURACEAE' =>  '绿片苔科 Aneuraceae H. Klinggr.',
'ANOMODONTACEAE' => '牛舌藓科 Anomodontaceae Kindb.',
'ANTHELIACEAE' =>  '兔耳苔科 Antheliaceae R. M. Schust.',
'ANTHOCEROTACEAE' =>  '角苔科 Anthocerotaceae Dumort.', 
'ANTITRICHIACEAE' => '逆毛藓科 Antitrichiaceae Ignatova & Ignatova',
'AONGSTROEMIACEAE' => '昂氏藓科 Aongstroemiaceae De Not.', 
'ARCHIDIACEAE' => '无轴藓科 Archidiaceae Schimp.',
'ARNELLIACEAE' =>  '阿氏苔科 Arnelliaceae Nakai',
'AULACOMNIACEAE' => '皱蒴藓科 Aulacomniaceae Schimp.',
'AYTONIACEAE' =>  '疣冠苔科 Aytoniaceae Cavers', 
'BALANTIOPSACEAE' => '小袋苔科 Balantiopsaceae H. Buch',
'BARTRAMIACEAE' => '珠藓科 Bartramiaceae Schw?gr.',
 'BLASIACEAE' => '壶苞苔科 Blasiaceae H. Klinggr.', 
 'BLEPHAROSTOMATACEAE' => '睫毛苔科 Blepharostomataceae W. Frey & M. Stech', 
 'BRACHYTHECIACEAE' => '青藓科 Brachytheciaceae Schimp.', 
 'BRUCHIACEAE' => '小烛藓科 Bruchiaceae Schimp.', 
 'BRYACEAE' => '真藓科 Bryaceae Schw?gr.',
 'BRYOWIJKIACEAE' => '蔓枝藓科 Bryowijkiaceae M. Stech & W. Frey', 
 'BRYOXIPHIACEAE' => '虾藓科 Bryoxiphiaceae Besch.', 
 'BUXBAUMIACEAE' => '烟杆藓科 Buxbaumiaceae Schimp.', 
 'CALLIERGONACEAE' => '湿原藓科 Calliergonaceae Vanderp. & Heden?s & C. J. Cox & A. J. Shaw',
 'CALYMPERACEAE' => '花叶藓科 Calymperaceae Kindb.', 
 'CALYPOGEIACEAE' => '护蒴苔科 Calypogeiaceae Arnell',
 'CEPHALOZIACEAE'  => '大萼苔科 Cephaloziaceae Mig.',
 'CEPHALOZIELLACEAE'  => '拟大萼苔科 Cephaloziellaceae Douin', 
 'CLAVEACEAE' => '星孔苔科 Claveaceae Cavers', 
 'CLIMACIACEAE' => '万年藓科 Climaciaceae Kindb.',
 'CONOCEPHALACEAE' => '蛇苔科 Conocephalaceae K. Müller ex Grolle',
 'CORSINIACEAE' => '花地钱科 Corsiniaceae Engl.',
 'CRYPHAEACEAE'  => '隐蒴藓科 Cryphaeaceae Schimp.',
 'CYATHODIACEAE' => '光苔科 Cyathodiaceae (Grolle) Stotler & Crand.-Stotl.', 
 'DALTONIACEAE' => '小黄藓科 Daltoniaceae Schimp.', 
 'DENDROCEROTACEAE' => '树角苔科 Dendrocerotaceae H?ssel', 
 'DICRANACEAE' => '曲尾藓科 Dicranaceae Schimp.', 
 'DICRANELLACEAE' => '小曲尾藓科 Dicranellaceae M. Stech',
 'DIPHYSCIACEAE' => '短颈藓科 Diphysciaceae M. Fleisch.', 
 'DITRICHACEAE' => '牛毛藓科 Ditrichaceae Limpr.',
 'DRUMMONDIACEAE' => '木衣藓科 Drummondiaceae Goffinet',
 'DUMORTIERACEAE' => '毛地钱科 Dumortieraceae D. G. Long',
 'ENCALYPTACEAE' => '大帽藓科 Encalyptaceae Schimp.', 
 'ENTODONTACEAE' => '绢藓科 Entodontaceae Kindb.',
 'EPHEMERACEAE' => '夭命藓科 Ephemeraceae Schimp.', 
 'ERPODIACEAE' => '树生藓科 Erpodiaceae Broth.',
 'EXORMOTHECACEAE' => '突苞苔科 Exormothecaceae Grolle',
 'FABRONIACEAE' => '碎米藓科 Fabroniaceae Schimp.', 
 'FISSIDENTACEAE' => '凤尾藓科 Fissidentaceae Schimp.',
 'FOLIOCEROTACEAE' => '褐角苔科 Foliocerotaceae H?ssel',
 'FONTINALACEAE'  => '水藓科 Fontinalaceae Schimp.', 
 'FOSSOMBRONIACEAE' => '小叶苔科 Fossombroniaceae Hazsl.',
 'FRULLANIACEAE' => '耳叶苔科 Frullaniaceae Lorch',
 'FUNARIACEAE' => '葫芦藓科 Funariaceae Schw?gr.',
 'GEOCALYCACEAE' => '地萼苔科 Geocalycaceae H. Klinggr.', 
 'GRIMMIACEAE' => '紫萼藓科 Grimmiaceae Arn.',
 'GYMNOMITRIACEAE' => '全萼苔科 Gymnomitriaceae H. Klinggr.',
 'HABRODONTACEAE' => '柔齿藓科 Habrodontaceae Schimp.',
 'HAPLOMITRIACEAE' => '裸蒴苔科 Haplomitriaceae Děde?ek', 
 'HEDWIGIACEAE'=> '虎尾藓科 Hedwigiaceae Schimp.', 
 'HERBERTACEAE'=> '剪叶苔科 Herbertaceae R. M. Schust.', 
 'HETEROCLADIACEAE' => '异枝藓科 Heterocladiaceae Ignatova & Ignatova',
 'HOOKERIACEAE' => '油藓科 Hookeriaceae Schimp.',
 'HYLOCOMIACEAE' => '塔藓科 Hylocomiaceae M. Fleisch.', 
 'HYPNACEAE' => '灰藓科 Hypnaceae Schimp.', 
 'HYPNODENDRACEAE' => '树灰藓科 Hypnodendraceae Broth.',
 'HYPOPTERYGIACEAE'  => '孔雀藓科 Hypopterygiaceae Mitt.',
 'JACKIELLACEAE' => '甲克苔科 Jackiellaceae R. M. Schust.',
 'JAMESONIELLACEAE' => '圆叶苔科 Jamesoniellaceae He-Nygrén & Julén & Ahonen & Glenny & Piippo', 
 'JUBULACEAE' => '毛耳苔科 Jubulaceae H. Klinggr.',
 'JUNGERMANNIACEAE' => '叶苔科 Jungermanniaceae Rchb.', 
 'LEJEUNEACEAE' => '细鳞苔科 Lejeuneaceae Cas.-Gil.',
 'LEMBOPHYLLACEAE'  => '船叶藓科 Lembophyllaceae Broth.',
 'LEPICOLEACEAE' => '复叉苔科 Lepicoleaceae R. M. Schust.',
 'LEPIDOLAENACEAE' => '多囊苔科 Lepidolaenaceae Nakai', 
 'LEPIDOZIACEAE' => '指叶苔科 Lepidoziaceae Nakai', 
 'LESKEACEAE' => '薄罗藓科 Leskeaceae Schimp.', 
 'LEUCOBRYACEAE' => '白发藓科 Leucobryaceae Schimp.', 
 'LEUCODONTACEAE' => '白齿藓科 Leucodontaceae Schimp.', 
 'LEUCOMIACEAE' => '白藓科 Leucomiaceae Broth.', 
 'LOPHOCOLEACEAE' => '齿萼苔科 Lophocoleaceae Vanden Berghen',
 'LOPHOZIACEAE'  => '裂叶苔科 Lophoziaceae Cavers', 
 'LUNULARIACEAE' => '半月苔科 Lunulariaceae H. Klinggr.', 
 'MAKINOACEAE' => '南溪苔科 Makinoaceae Giacom.', 
 'MARCHANTIACEAE' => '地钱科 Marchantiaceae Lindl.', 
 'MASTIGOPHORACEAE' => '须苔科 Mastigophoraceae R. M. Schust.',
 'MEESIACEAE'  => '寒藓科 Meesiaceae Schimp.',
 'METEORIACEAE' => '蔓藓科 Meteoriaceae Kindb.',
 'METZGERIACEAE'  => '叉苔科 Metzgeriaceae H. Klinggr.', 
 'MNIACEAE' => '提灯藓科 Mniaceae Schw?gr.',
 'MOERCKIACEAE'  => '莫氏苔科 Moerckiaceae Stotler & Crand.-Stotl.',
 'MONOSOLENIACEAE'  => '单月苔科 Monosoleniaceae E. H. Wilson', 
 'MYLIACEAE' => '小萼苔科 Myliaceae (Grolle) Schljakov',
 'MYURIACEAE' => '金毛藓科 Myuriaceae M. Fleisch.',
 'NECKERACEAE'  => '平藓科 Neckeraceae Schimp.',
 'NEOTRICHOCOLEACEAE'  => '新绒苔科 Neotrichocoleaceae Inoue', 
 'NOTOTHYLADACEAE' => '短角苔科 Notothyladaceae Müll. Frib. ex Prosk.', 
 'OEDIPODIACEAE' => '长台藓科 Oedipodiaceae Schimp.',
 'ONCOPHORACEAE'  => '曲背藓科 Oncophoraceae M. Stech',
 'ORTHODONTIACEAE' => '直齿藓科 Orthodontiaceae Goffinet',
 'ORTHOTRICHACEAE' => '木灵藓科 Orthotrichaceae Arn.', 
 'PALLAVICINIACEAE' => '带叶苔科 Pallaviciniaceae Mig.', 
 'PELLIACEAE' => '溪苔科 Pelliaceae H. Klinggr.', 
 'PILOTRICHACEAE' => '茸帽藓科 Pilotrichaceae Kindb.', 
 'PLAGIOCHILACEAE' => '羽苔科 Plagiochilaceae Müll. Frib.',
 'PLAGIOTHECIACEAE'  => '棉藓科 Plagiotheciaceae M. Fleisch.',
 'PLEUROZIACEAE'  => '紫叶苔科 Pleuroziaceae Müll. Frib.',
 'POLYTRICHACEAE' => '金发藓科 Polytrichaceae Schw?gr.',
 'PORELLACEAE' => '光萼苔科 Porellaceae Cavers', 
 'POTTIACEAE' => '丛藓科 Pottiaceae Schimp.', 
 'PSEUDOLEPICOLEACEAE' => '拟复叉苔科 Pseudolepicoleaceae Fulford & J. Taylor',
 'PSEUDOLESKEACEAE' => '拟薄罗藓科 Pseudoleskeaceae Schimp.', 
 'PSEUDOLESKEELLACEAE' => '假细罗藓科 Pseudoleskeellaceae Ignatova & Ignatova', 
 'PTERIGYNANDRACEAE' => '腋苞藓科 Pterigynandraceae Schimp.', 
 'PTEROBRYACEAE' => '蕨藓科 Pterobryaceae Kindb.', 
 'PTILIDIACEAE' => '毛叶苔科 Ptilidiaceae H. Klinggr.', 
 'PTYCHOMITRIACEAE' => '缩叶藓科 Ptychomitriaceae Schimp.',
 'PTYCHOMNIACEAE' => '棱蒴藓科 Ptychomniaceae M. Fleisch.',
 'PYLAISIACEAE' => '金灰藓科 Pylaisiaceae Schimp.',
 'PYLAISIADELPHACEAE' => '毛锦藓科 Pylaisiadelphaceae Goffinet & W. R. Buck',
 'RACOPILACEAE' => '卷柏藓科 Racopilaceae Kindb.', 
 'RADULACEAE' => '扁萼苔科 Radulaceae (Dumort.) K. Müller', 
 'REGMATODONTACEAE' => '异齿藓科 Regmatodontaceae Broth.',
 'RHACHITHECIACEAE' => '刺藓科 Rhachitheciaceae H. Rob.',
 'RHIZOGONIACEAE' => '桧藓科 Rhizogoniaceae Broth.', 
 'RHYTIDIACEAE' => '垂枝藓科 Rhytidiaceae Broth.', 
 'RICCIACEAE' => '钱苔科 Ricciaceae Rchb.',
 'SCAPANIACEAE' => '合叶苔科 Scapaniaceae Mig.',
 'SCHISTOCHILACEAE' => '歧舌苔科 Schistochilaceae H. Buch',
 'SCHISTOSTEGACEAE' => '光藓科 Schistostegaceae Schimp.',
 'SCORPIDIACEAE' => '蝎尾藓科 Scorpidiaceae Ignatova & Ignatova',
 'SELIGERIACEAE' => '细叶藓科 Seligeriaceae Schimp.',
 'SEMATOPHYLLACEAE' => '锦藓科 Sematophyllaceae Broth.',
 'SPHAEROCARPACEAE' => '囊果苔科 Sphaerocarpaceae',
 'SPHAGNACEAE' => '泥炭藓科 Sphagnaceae Dumort.', 
 'SPLACHNACEAE' => '壶藓科 Splachnaceae Grev. & Arn.',
 'STEREOPHYLLACEAE' => '硬叶藓科 Stereophyllaceae W. R. Buck & R. R. Ireland', 
 'SYMPHYODONTACEAE' => '刺果藓科 Symphyodontaceae M. Fleisch.', 
 'TAKAKIACEAE' => '藻苔科 Takakiaceae M. Stech & W. Frey',
 'TARGIONIACEAE'  => '皮叶苔科 Targioniaceae Dumort.',
 'TETRAPHIDACEAE' => '四齿藓科 Tetraphidaceae Schimp.', 
 'THUIDIACEAE' => '羽藓科 Thuidiaceae Schimp.', 
 'TIMMIACEAE' => '美姿藓科 Timmiaceae Schimp.', 
 'TRACHYLOMATACEAE' => '粗柄藓科 Trachylomataceae (M. Fleisch.) W. R. Buck & Vitt.',
 'TREUBIACEAE' => '陶氏苔科 Treubiaceae Verd.',
 'TRICHOCOLEACEAE' => '绒苔科 Trichocoleaceae Nakai',
 'WIESNERELLACEAE' => '魏氏苔科 Wiesnerellaceae Inoue' 
)  
  w.save(working_file)
  puts"DEBUG: complete"
end
