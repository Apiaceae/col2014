# encoding: utf-8

def need_sort(s)
  sorted_province = sort_province(array_province(s))
  origin_province =  dis_province(s)

  if origin_province != sorted_province
    return true
  else
    return false
  end
end

def dis_arbord(dis_string)
  if dis_string.include? ';'
    pos_of_semicolon = dis_string.index(';')
    return dis_string[0, pos_of_semicolon].lstrip
  end
end

def dis_province(dis_string)
  if dis_string.include? ';'
    pos_of_semicolon = dis_string.index(';')
    dis_string_aboard = dis_string[0, pos_of_semicolon]
    return dis_string[pos_of_semicolon+1, dis_string.length - dis_string_aboard.length].lstrip
  else
    return dis_string
  end
end

def array_province(dis_string)
  if dis_string.include? ';'
    pos_of_semicolon = dis_string.index(';')
    dis_string_aboard = dis_string[0, pos_of_semicolon]
    dis_string_china = dis_string[pos_of_semicolon+1, dis_string.length - dis_string_aboard.length+1].lstrip
    return dis_string_china.split(',')
  end
  if !dis_string.include? ';'
    return dis_string.split(',')
  end
end

def sort_province(array)
  dis_string_china_ordered = {}
  array.each do |pv|
    if pv == "黑龙江"
      dis_string_china_ordered["黑龙江"] = 1
    end
    if pv == "吉林"
      dis_string_china_ordered["吉林"] = 2
    end
    if pv == "辽宁"
      dis_string_china_ordered["辽宁"] = 3
    end
    if pv == "内蒙古"
      dis_string_china_ordered["内蒙古"] = 4
    end
    if pv == "河北"
      dis_string_china_ordered["河北"] = 5
    end
    if pv == "北京"
      dis_string_china_ordered["北京"] = 6
    end
    if pv == "天津"
      dis_string_china_ordered["天津"] = 7
    end
    if pv == "山西"
      dis_string_china_ordered["山西"] = 8
    end
    if pv == "山东"
      dis_string_china_ordered["山东"] = 9
    end
    if pv == "河南"
      dis_string_china_ordered["河南"] = 10
    end
    if pv == "陕西"
      dis_string_china_ordered["陕西"] = 11
    end
    if pv == "宁夏"
      dis_string_china_ordered["宁夏"] = 12
    end
    if pv == "甘肃"
      dis_string_china_ordered["甘肃"] = 13
    end
    if pv == "青海"
      dis_string_china_ordered["青海"] = 14
    end
    if pv == "新疆"
      dis_string_china_ordered["新疆"] = 15
    end
    if pv == "安徽"
      dis_string_china_ordered["安徽"] = 16
    end
    if pv == "江苏"
      dis_string_china_ordered["江苏"] = 17
    end
    if pv == "上海"
      dis_string_china_ordered["上海"] = 18
    end
    if pv == "浙江"
      dis_string_china_ordered["浙江"] = 19
    end
    if pv == "江西"
      dis_string_china_ordered["江西"] = 20
    end
    if pv == "湖南"
      dis_string_china_ordered["湖南"] = 21
    end
    if pv == "湖北"
      dis_string_china_ordered["湖北"] = 22
    end
    if pv == "四川"
      dis_string_china_ordered["四川"] = 23
    end
    if pv == "重庆"
      dis_string_china_ordered["重庆"] = 24
    end
    if pv == "贵州"
      dis_string_china_ordered["贵州"] = 25
    end
    if pv == "云南"
      dis_string_china_ordered["云南"] = 26
    end
    if pv == "西藏"
      dis_string_china_ordered["西藏"] = 27
    end
    if pv == "福建"
      dis_string_china_ordered["福建"] = 28
    end
    if pv == "台湾"
      dis_string_china_ordered["台湾"] = 29
    end
    if pv == "广东"
      dis_string_china_ordered["广东"] = 30
    end
    if pv == "广西"
      dis_string_china_ordered["广西"] = 31
    end
    if pv == "海南"
      dis_string_china_ordered["海南"] = 32
    end
    if pv == "香港"
      dis_string_china_ordered["香港"] = 33
    end
    if pv == "澳门"
      dis_string_china_ordered["澳门"] = 34
    end
    # termporarly treat un standard province statements to the end of the distribution
    if pv.length > 3
      dis_string_china_ordered["#{pv}"] = 35
    end
  end
  ordered_string = ""
  ordered_hash = Hash[dis_string_china_ordered.invert.sort]
  ordered_hash.each do |k, v|
    ordered_string << v+","
  end
  return ordered_string.chop
end
