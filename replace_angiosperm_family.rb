#!/usr/bin/env ruby
# encoding: utf-8

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
  working_file = ARGV[1] || file.sub(/\.docx/, '-replaced.docx')
  w = WordXmlManipulate.open(file)
  w.merge('ACANTHACEAE' => '爵床科 Acanthaceae Juss.',
'ACHARIACEAE' => '钟花科 Achariaceae Harms',
'ACORACEAE' => '菖蒲科 Acoraceae L.',
'ACTINIDIACEAE' => '猕猴桃科 Actinidiaceae Engl. & Gilg',
'ADOXACEAE' => '五福花科 Adoxaceae E. Mey.',
'AIZOACEAE' => '番杏科 Aizoaceae Martinov',
'AKANIACEAE' => '叠珠树科 Akaniaceae Stapf',
'ALISMATACEAE' => '泽泻科 Alismataceae Vent.',
'ALTINGIACEAE' => '枫香科 Altingiaceae Horan.',
'AMARANTHACEAE' => '苋科 Amaranthaceae Juss.',
'AMARYLLIDACEAE' => '石蒜科 Amaryllidaceae J. St.-Hil.',
'ANACARDIACEAE' => '漆树科 Anacardiaceae R. Br.',
'ANCISTROCLADACEAE' => '钩枝藤科 Ancistrocladaceae Planch. ex Walp.',
'ANNONACEAE' => '番荔枝科 Annonaceae Juss.',
'APIACEAE' => '伞形科 Apiaceae Lindl.',
'APOCYNACEAE' => '夹竹桃科 Apocynaceae Juss.',
'APONOGETONACEAE' => '水蕹科 Aponogetonaceae Planch.',
'AQUIFOLIACEAE' => '冬青科 Aquifoliaceae Bercht. & J. Presl',
'ARACEAE' => '天南星科 Araceae Juss.',
'ARALIACEAE' => '五加科 Araliaceae Juss.',
'ARECACEAE' => '棕榈科 Arecaceae Bercht. & J. Presl',
'ARISTOLOCHIACEAE' => '马兜铃科 Aristolochiaceae Juss.',
'ASPARAGACEAE' => '天门冬科 Asparagaceae Juss.',
'ASTERACEAE' => '菊科 Asteraceae Martynov',
'BALANOPHORACEAE' => '蛇菰科 Balanophoraceae Benth. & Hook. f.',
'BALSAMINACEAE' => '凤仙花科 Balsaminaceae Bercht. & J. Presl',
'BASELLACEAE' => '落葵科 Basellaceae Durandé',
'BEGONIACEAE' => '秋海棠科 Begoniaceae Bercht. & J. Presl',
'BERBERIDACEAE' => '小檗科 Berberidaceae Durandé',
'BETULACEAE' => '桦木科 Betulaceae Gray',
'BIEBERSTEINIACEAE' => '熏倒牛科 Biebersteiniaceae Endl.',
'BIGNONIACEAE' => '紫葳科 Bignoniaceae Durandé',
'BIXACEAE' => '红木科 Bixaceae Link',
'BORAGINACEAE' => '紫草科 Boraginaceae Adans.',
'BORTHWICKIACEAE' => '节蒴木科 Borthwickiaceae J.X. Su & Wei Wang & Li Bing Zhang & Z.D.Chen',
'BRASSICACEAE' => '十字花科 Brassicaceae Burnett',
'BROMELIACEAE' => '凤梨科 Bromeliaceae Juss.',
'BURMANNIACEAE' => '水玉簪科 Burmanniaceae Blume',
'BURSERACEAE' => '橄榄科 Burseraceae Kunth',
'BUTOMACEAE' => '花蔺科 Butomaceae Mirb.',
'BUXACEAE' => '黄杨科 Buxaceae Dumort.',
'CABOMBACEAE' => '莼菜科 Cabombaceae Rich. ex A. Rich.',
'CACTACEAE' => '仙人掌科 Cactaceae Durandé',
'CALOPHYLLACEAE' => '红厚壳科 Calophyllaceae J. Agardh',
'CALYCANTHACEAE' => '腊梅科 Calycanthaceae Lindl.',
'CAMPANULACEAE' => '桔梗科 Campanulaceae Adans.',
'CANNABACEAE' => '大麻科 Cannabaceae Martynov',
'CANNACEAE' => '美人蕉科 Cannaceae Durandé',
'CAPPARACEAE' => '山柑科 Capparaceae Adans.',
'CAPRIFOLIACEAE' => '忍冬科 Caprifoliaceae Adans.',
'CARDIOPTERIDACEAE' => '心翼果科 Cardiopteridaceae Blume',
'CARICACEAE' => '番木瓜科 Caricaceae Dumort.',
'CARLEMANNIACEAE' => '香茜科 Carlemanniaceae Airy Shaw',
'CARYOPHYLLACEAE' => '石竹科 Caryophyllaceae Durandé',
'CASUARINACEAE' => '木麻黄科 Casuarinaceae R. Br.',
'CELASTRACEAE' => '卫矛科 Celastraceae R. Br.',
'CENTROLEPIDACEAE' => '刺鳞草科 Centrolepidaceae Endl.',
'CENTROPLACACEAE' => '扁距木科 Centroplacaceae Doweld & Reveal',
'CERATOPHYLLACEAE' => '金鱼藻科 Ceratophyllaceae Gray',
'CERCIDIPHYLLACEAE' => '连香树科 Cercidiphyllaceae Engl.',
'CHLORANTHACEAE' => '金粟兰科 Chloranthaceae R. Br.',
'CIRCAEASTERACEAE' => '星叶草科 Circaeasteraceae Kuntze ex Hutch.',
'CISTACEAE' => '半日花科 Cistaceae Adans.',
'CLEOMACEAE' => '白花菜科 Cleomaceae Horan.',
'CLETHRACEAE' => '桤叶树科 Clethraceae Klotzsch',
'CLUSIACEAE' => '山竹子科(藤黄科) Clusiaceae Lindl.',
'COLCHICACEAE' => '秋水仙科 Colchicaceae DC.',
'COMBRETACEAE' => '使君子科 Combretaceae R. Br.',
'COMMELINACEAE' => '鸭跖草科 Commelinaceae Mirb.',
'CONNARACEAE' => '牛栓藤科 Connaraceae R. Br.',
'CONVOLVULACEAE' => '旋花科 Convolvulaceae Durandé',
'CORIARIACEAE' => '马桑科 Coriariaceae DC.',
'CORNACEAE' => '山茱萸科 Cornaceae Dumort.',
'CORSIACEAE' => '白玉簪科 Corsiaceae Becc.',
'COSTACEAE' => '闭鞘姜科 Costaceae Nakai',
'CRASSULACEAE' => '景天科 Crassulaceae J. St.-Hil.',
'CRYPTERONIACEAE' => '隐翼科 Crypteroniaceae A. DC.',
'CUCURBITACEAE' => '葫芦科 Cucurbitaceae Durandé',
'CYMODOCEACEAE' => '丝粉藻科 Cymodoceaceae Vines',
'CYNOMORIACEAE' => '锁阳科 Cynomoriaceae Lindl.',
'CYPERACEAE' => '莎草科 Cyperaceae Juss.',
'DAPHNIPHYLLACEAE' => '交让木科 Daphniphyllaceae Müll.Arg.',
'DIAPENSIACEAE' => '岩梅科 Diapensiaceae Lindl.',
'DICHAPETALACEAE' => '毒鼠子科 Dichapetalaceae Baill.',
'DILLENIACEAE' => '五桠果科 Dilleniaceae Salisb.',
'DIOSCOREACEAE' => '薯蓣科 Dioscoreaceae R. Br.',
'DIPENTODONTACEAE' => '十齿花科 Dipentodontaceae Merr.',
'DIPTEROCARPACEAE' => '龙脑香科 Dipterocarpaceae Blume',
'DROSERACEAE' => '茅膏菜科 Droseraceae Salisb.',
'EBENACEAE' => '柿树科 Ebenaceae Gürcke',
'ELAEAGNACEAE' => '胡颓子科 Elaeagnaceae Adans.',
'ELAEOCARPACEAE' => '杜英科 Elaeocarpaceae Juss. ex DC.',
'ELATINACEAE' => '沟繁缕科 Elatinaceae Dumort.',
'ERICACEAE' => '杜鹃花科 Ericaceae Durandé',
'ERIOCAULACEAE' => '谷精草科 Eriocaulaceae Martynov',
'ERYTHROXYLACEAE' => '古柯科 Erythroxylaceae Kunth',
'ESCALLONIACEAE' => '南鼠刺科 Escalloniaceae R.Brown. ex Dumort.',
'EUCOMMIACEAE' => '杜仲科 Eucommiaceae Van Tiegh.',
'EUPHORBIACEAE' => '大戟科 Euphorbiaceae J. F. Gmel.',
'EUPTELEACEAE' => '领春木科 Eupteleaceae K. Wilh.',
'FABACEAE' => '豆科 Fabaceae Lindl.',
'FAGACEAE' => '壳斗科 Fagaceae Dumort.',
'FLAGELLARIACEAE' => '须叶藤科 Flagellariaceae Dum.',
'FRANKENIACEAE' => '瓣鳞花科 Frankeniaceae S. F. Gray',
'GARRYACEAE' => '丝缨花科 Garryaceae Lindl.',
'GELSEMIACEAE' => '钩吻科 Gelsemiaceae Struwe & V. A. Albert',
'GENTIANACEAE' => '龙胆科 Gentianaceae Durandé',
'GERANIACEAE' => '牻牛儿苗科 Geraniaceae Adans.',
'GESNERIACEAE' => '苦苣苔科 Gesneriaceae Rich. & Juss. ex DC.',
'GISEKIACEAE' => '晶粟草科 Gisekiaceae Nakai',
'GOODENIACEAE' => '草海桐科 Goodeniaceae R. Br.',
'GROSSULARIACEAE' => '茶藨子科 Grossulariaceae DC.',
'HALORAGACEAE' => '小二仙草科 Haloragaceae R. Br.',
'HAMAMELIDACEAE' => '金缕梅科 Hamamelidaceae R. Br.',
'HELWINGIACEAE' => '青荚叶科 Helwingiaceae Decne.',
'HERNANDIACEAE' => '莲叶桐科 Hernandiaceae Bercht. & J. Presl',
'HYDRANGEACEAE' => '绣球花科 Hydrangeaceae Dumort.',
'HYDROCHARITACEAE' => '水鳖科 Hydrocharitaceae Juss.',
'HYDROPHYLLACEAE' => '田基麻科 Hydrophyllaceae R. Br.',
'HYPERICACEAE' => '金丝桃科 Hypericaceae Juss.',
'HYPOXIDACEAE' => '小金梅草科 Hypoxidaceae R.Brown.',
'ICACINACEAE' => '茶茱萸科 Icacinaceae (Benth.) Miers',
'IRIDACEAE' => '鸢尾科 Iridaceae Durandé',
'ITEACEAE' => '鼠刺科 Iteaceae J. Agardh',
'IXIOLIRIACEAE' => '鸢尾蒜科 Ixioliriaceae Nakai',
'IXONANTHACEAE' => '黏木科 Ixonanthaceae Planch. ex Miq.',
'JUGLANDACEAE' => '胡桃科 Juglandaceae DC. ex Perleb',
'JUNCACEAE' => '灯心草科 Juncaceae Durandé',
'JUNCAGINACEAE' => '水麦冬科 Juncaginaceae Rich.',
'LAMIACEAE' => '唇形科 Lamiaceae Martynov',
'LARDIZABALACEAE' => '木通科 Lardizabalaceae R. Br.',
'LAURACEAE' => '樟科 Lauraceae Durandé',
'LECYTHIDACEAE' => '玉蕊科 Lecythidaceae A. Rich.',
'LENTIBULARIACEAE' => '狸藻科 Lentibulariaceae Rich.',
'LILIACEAE' => '百合科 Liliaceae Adans.',
'LINACEAE' => '亚麻科 Linaceae DC. ex Perleb',
'LINDERNIACEAE' => '母草科 Linderniaceae Borsch & K. Müller & Eb. Fisch.',
'LOGANIACEAE' => '马钱科 Loganiaceae R. Br.',
'LORANTHACEAE' => '桑寄生科 Loranthaceae Juss.',
'LOWIACEAE' => '兰花蕉科 Lowiaceae Ridl.',
'LYTHRACEAE' => '千屈菜科 Lythraceae J. St.-Hil.',
'MAGNOLIACEAE' => '木兰科 Magnoliaceae Juss.',
'MALPIGHIACEAE' => '金虎尾科 Malpighiaceae Durandé',
'MALVACEAE' => '锦葵科 Malvaceae Adans.',
'MARANTACEAE' => '竹芋科 Marantaceae R. Br.',
'MARTYNIACEAE' => '角胡麻科 Martyniaceae Horan.',
'MASTIXIACEAE' => '单室茱萸科 Mastixiaceae Calestani',
'MELANTHIACEAE' => '黑药花科 Melanthiaceae Batsch ex Borkh.',
'MELASTOMATACEAE' => '野牡丹科 Melastomataceae Juss.',
'MELIACEAE' => '楝科 Meliaceae Juss.',
'MENISPERMACEAE' => '防己科 Menispermaceae Juss.',
'MENYANTHACEAE' => '睡菜科 Menyanthaceae Bercht. & J. Presl',
'MITRASTEMONACEAE' => '帽蕊草科 Mitrastemonaceae Makino',
'MOLLUGINACEAE' => '粟米草科 Molluginaceae Raf.',
'MORACEAE' => '桑科 Moraceae Link',
'MORINGACEAE' => '辣木科 Moringaceae Martynov',
'MUSACEAE' => '芭蕉科 Musaceae Durandé',
'MYOPORACEAE' => '苦槛蓝科 Myoporaceae R. Br.',
'MYRICACEAE' => '杨梅科 Myricaceae A. Rich.',
'MYRISTICACEAE' => '肉豆蔻科 Myristicaceae R. Br.',
'MYRTACEAE' => '桃金娘科 Myrtaceae Adans.',
'NARTHECIACEAE' => '纳茜菜科 Nartheciaceae Fr. ex Bjurzon',
'NELUMBONACEAE' => '莲科 Nelumbonaceae Bercht. & J. Presl',
'NEPENTHACEAE' => '猪笼草科 Nepenthaceae Bercht. & J. Presl',
'NITRARIACEAE' => '白刺科 Nitrariaceae Berch. & J. Presl',
'NYCTAGINACEAE' => '紫茉莉科 Nyctaginaceae Juss.',
'NYMPHAEACEAE' => '睡莲科 Nymphaeaceae Salisb.',
'OCHNACEAE' => '金莲木科 Ochnaceae DC.',
'OLACACEAE' => '铁青树科 Olacaceae Juss. ex R. Br.',
'OLEACEAE' => '木犀科 Oleaceae Hoffmanns. & Link',
'ONAGRACEAE' => '柳叶菜科 Onagraceae Adans.',
'OPILIACEAE' => '山柚子科 Opiliaceae Valeton',
'ORCHIDACEAE' => '兰科 Orchidaceae Adans.',
'OROBANCHACEAE' => '列当科 Orobanchaceae Vent.',
'OXALIDACEAE' => '酢浆草科 Oxalidaceae R. Br.',
'PAEONIACEAE' => '芍药科 Paeoniaceae Raf.',
'PANDACEAE' => '小盘木科 Pandaceae Engl. & Gilg',
'PANDANACEAE' => '露兜树科 Pandanaceae R. Br.',
'PAPAVERACEAE' => '罂粟科 Papaveraceae Adans.',
'PASSIFLORACEAE' => '西番莲科 Passifloraceae Juss. ex Roussel',
'PAULOWNIACEAE' => '泡桐科 Paulowniaceae Nakai',
'PEDALIACEAE' => '胡麻科 Pedaliaceae R. Br.',
'PENTAPHRAGMATACEAE' => '五膜草科 Pentaphragmataceae J.G. Agardh',
'PENTAPHYLACACEAE' => '五列木科 Pentaphylacaceae Engl.',
'PENTHORACEAE' => '扯根菜科 Penthoraceae Rydb. ex Britt.',
'PETROSAVIACEAE' => '无叶莲科 Petrosaviaceae Hutch.',
'PHILYDRACEAE' => '田葱科 Philydraceae Link',
'PHRYMACEAE' => '透骨草科 Phrymaceae Schauer',
'PHYLLANTHACEAE' => '叶下珠科 Phyllanthaceae Martinov',
'PHYTOLACCACEAE' => '商陆科 Phytolaccaceae Durandé',
'PIPERACEAE' => '胡椒科 Piperaceae Batsch',
'PITTOSPORACEAE' => '海桐花科 Pittosporaceae R. Br.',
'PLANTAGINACEAE' => '车前科 Plantaginaceae Durandé',
'PLATANACEAE' => '悬铃木科 Platanaceae T. Lestib.',
'PLUMBAGINACEAE' => '白花丹科 Plumbaginaceae Durandé',
'POACEAE' => '禾本科 Poaceae Caruel',
'PODOSTEMACEAE' => '川苔草科 Podostemaceae Kunth',
'POLEMONIACEAE' => '花荵科 Polemoniaceae Juss.',
'POLYGALACEAE' => '远志科 Polygalaceae Hoffmanns. & Link',
'POLYGONACEAE' => '蓼科 Polygonaceae Durandé',
'PONTEDERIACEAE' => '雨久花科 Pontederiaceae Kunth',
'PORTULACACEAE' => '马齿苋科 Portulacaceae Adans.',
'POSIDONIACEAE' => '波喜荡科 Posidoniaceae Vines',
'POTAMOGETONACEAE' => '眼子菜科 Potamogetonaceae Rchb.',
'PRIMULACEAE' => '报春花科 Primulaceae Batsch',
'PROTEACEAE' => '山龙眼科 Proteaceae Juss.',
'PUTRANJIVACEAE' => '核果木科 Putranjivaceae Meisn.',
'RAFFLESIACEAE' => '大花草科 Rafflesiaceae Dumort.',
'RANUNCULACEAE' => '毛茛科 Ranunculaceae Adans.',
'RESEDACEAE' => '木犀草科 Resedaceae Bercht. & J. Presl',
'RESTIONACEAE' => '帚灯草科 Restionaceae R. Br.',
'RHAMNACEAE' => '鼠李科 Rhamnaceae Durandé',
'RHIZOPHORACEAE' => '红树科 Rhizophoraceae Pers.',
'ROSACEAE' => '蔷薇科 Rosaceae Adans.',
'RUBIACEAE' => '茜草科 Rubiaceae Durandé',
'RUPPIACEAE' => '川蔓藻科 Ruppiaceae Horn.',
'RUTACEAE' => '芸香科 Rutaceae Durandé',
'SABIACEAE' => '清风藤科 Sabiaceae Blume',
'SALICACEAE' => '杨柳科 Salicaceae Mirb.',
'SALVADORACEAE' => '刺茉莉科 Salvadoraceae Lindl.',
'SANTALACEAE' => '檀香科 Santalaceae R. Br.',
'SAPINDACEAE' => '无患子科 Sapindaceae Juss.',
'SAPOTACEAE' => '山榄科 Sapotaceae Durandé',
'SAURURACEAE' => '三白草科 Saururaceae Martynov',
'SAXIFRAGACEAE' => '虎耳草科 Saxifragaceae Durandé',
'SCHEUCHZERIACEAE' => '冰沼草科 Scheuchzeriaceae F. Rudolphi',
'SCHISANDRACEAE' => '五味子科 Schisandraceae Blume',
'SCHOEPFIACEAE' => '青皮木科 Schoepfiaceae Blume',
'SCROPHULARIACEAE' => '玄参科 Scrophulariaceae Durandé',
'SIMAROUBACEAE' => '苦木科 Simaroubaceae DC.',
'SLADENIACEAE' => '肋果茶科 Sladeniaceae Airy Shaw',
'SMILACACEAE' => '菝葜科 Smilacaceae Vent.',
'SOLANACEAE' => '茄科 Solanaceae Adans.',
'SPHENOCLEACEAE' => '尖瓣花科 Sphenocleaceae T. Baskerv.',
'STACHYURACEAE' => '旌节花科 Stachyuraceae J. Agardh',
'STAPHYLEACEAE' => '省沽油科 Staphyleaceae Martynov',
'STEMONACEAE' => '百部科 Stemonaceae Caruel',
'STEMONURACEAE' => '粗丝木科(金檀木科) Stemonuraceae K?rehed',
'STRELITZIACEAE' => '梧桐科 Sterculiaceae Vent. ex Salib.',
'STYLIDIACEAE' => '花柱草科 Stylidiaceae R. Br.',
'STYRACACEAE' => '安息香科 Styracaceae DC. & Spreng.',
'SURIANACEAE' => '海人树科 Surianaceae Arn.',
'SYMPLOCACEAE' => '山矾科 Symplocaceae Desf.',
'TACCACEAE' => '箭根薯科 Taccaceae Bercht. & J. Presl',
'TALINACEAE' => '土人参科 Talinaceae Doweld',
'TAMARICACEAE' => '柽柳科 Tamaricaceae Bercht. & J. Presl',
'TAPISCIACEAE' => '瘿椒树科 Tapisciaceae Takht.',
'TETRAMELACEAE' => '四数木科 Tetramelaceae Airy Shaw',
'THEACEAE' => '山茶科 Theaceae Ker Gawl.',
'THELIGONACEAE' => '牛繁缕科 Theligonaceae Dumort.',
'THYMELAEACEAE' => '瑞香科 Thymelaeaceae Adans.',
'TILIACEAE' => '椴树科 Tiliaceae Adans.',
'TOFIELDIACEAE' => '岩菖蒲科 Tofieldiaceae Takht.',
'TORICELLIACEAE' => '鞘柄木科 Toricelliaceae (Wang) Hu',
'TRIURIDACEAE' => '霉草科 Triuridaceae Gardner',
'TROCHODENDRACEAE' => '昆栏树科 Trochodendraceae Prantl',
'TROPAEOLACEAE' => '旱金莲科 Tropaeolaceae DC.',
'TYPHACEAE' => '香蒲科 Typhaceae Durandé',
'ULMACEAE' => '榆科 Ulmaceae Mirb.',
'URTICACEAE' => '荨麻科 Urticaceae Durandé',
'VELLOZIACEAE' => '翡若翠科 Velloziaceae J. Agardh',
'VERBENACEAE' => '马鞭草科 Verbenaceae Adans.',
'VIOLACEAE' => '堇菜科 Violaceae Batsch',
'VITACEAE' => '葡萄科 Vitaceae Durandé',
'XANTHORRHOEACEAE' => '黄脂木科 Xanthorrhoeaceae Dumort.',
'XYRIDACEAE' => '黄眼草科 Xyridaceae C. Agardh',
'ZINGIBERACEAE' => '姜科 Zingiberaceae Adans.',
'ZOSTERACEAE' => '大叶藻科 Zosteraceae Dumort.',
'ZYGOPHYLLACEAE' => '蒺藜科 Zygophyllaceae R. Br.'
)
  w.save(working_file)
  puts"DEBUG: complete"
end
