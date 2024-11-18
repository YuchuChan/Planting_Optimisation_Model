library("readxl")
library("forecast") 
library(writexl)
library(showtext) 
showtext_auto()
library(dplyr) 
library("data.table") 

a = data.table(read_excel(path = "附件1.xlsx", sheet = "乡村的现有耕地"))
b = data.table(read_excel(path = "附件1.xlsx", sheet = "乡村种植的农作物"))
c = data.table(read_excel(path = "附件2.xlsx", sheet = "2023年的农作物种植情况"))
d = data.table(read_excel(path = "附件2.xlsx", sheet = "2023年统计的相关数据")) 

a = a[,-4]
# 使用gsub函数提取字母和数字
a$地块类别 <- gsub("[^A-Za-z]", "", a$地块名称)  # 移除非字母字符，保留字母
a$地块序号 <- gsub("[^0-9]", "", a$地块名称)  # 移除非数字字符，保留数字
a = (a[,-1])[,c(3,4,1,2)]

b = b[-c(42,43,44,45),]
name_var = b[,c(2,3)]
setkey(name_var,作物名称)

for(i in 1:nrow(c)){
  if(is.na(c[i,1])){
    c[i,1] = c[i-1,1]
  }
}
write_xlsx(c,path="2023年种植方案.xlsx")

c$地块类别 <- gsub("[^A-Za-z]", "", c$种植地块)  # 移除非字母字符，保留字母
c$地块序号 <- gsub("[^0-9]", "", c$种植地块)  # 移除非数字字符，保留数字
c = (c[,-1])[,c(6,7,1,2,3,4,5)]

d = d[-c(108,109,110),]

smart_peng_1q = d[d$地块类型 == '普通大棚']
smart_peng_1q = data.table(smart_peng_1q[smart_peng_1q$种植季次 == '第一季'])
for(i in 1:nrow(smart_peng_1q)){
  smart_peng_1q[i,4] = '智慧大棚'
}
d_total = rbind(d,smart_peng_1q)
d_total$地块类别 = d_total$地块类型

for(i in 1:nrow(d_total)){
  if(d_total[i,4] == '平旱地'){
    d_total[i,9] = 'A'
  }else if(d_total[i,4] == '梯田'){
    d_total[i,9] = 'B'
  }else if(d_total[i,4] == '山坡地'){
    d_total[i,9] = 'C'
  }else if(d_total[i,4] == '水浇地'){
    d_total[i,9] = 'D'
  }else if(d_total[i,4] == '普通大棚'){
    d_total[i,9] = 'E'
  }else if(d_total[i,4] == '智慧大棚'){
    d_total[i,9] = 'F'
  }
}

plant_S = data.table(aggregate((`种植面积/亩`)~(地块类别+作物名称+种植季次),data=c,FUN = sum))
colnames(plant_S)[4] <- '种植面积/亩'
setkey(plant_S,地块类别,作物名称,种植季次)

setkey(d_total,地块类别,作物名称,种植季次)

d_m = plant_S[d_total]
d_m = d_m[!which( is.na (d_m$`种植面积/亩`))]
d_m$`产量(斤)` = d_m$`亩产量/斤` * d_m$`种植面积/亩`
d_m$`成本(元)` = d_m$`种植成本/(元/亩)` * d_m$`种植面积/亩`

# 使用strsplit函数拆分字符串，并使用as.numeric转换为浮点数
split_ranges <- strsplit(as.character(d_m$`销售单价/(元/斤)`), "-")
d_m$`单价下限(元/斤)` <- as.numeric(sapply(split_ranges, `[`, 1))
d_m$`单价上限(元/斤)` <- as.numeric(sapply(split_ranges, `[`, 2))
d_m$`平均单价(元/斤)` <- (d_m$`单价下限(元/斤)`+d_m$`单价上限(元/斤)`)/2

d_m$序号 = as.numeric(d_m$序号)
d_m$作物编号 = as.numeric(d_m$作物编号)
setkey(d_m,作物名称)
d_m = d_m[name_var]
d_m = d_m[,c(5,6,2,16,1,7,3,8,9,10,4,11,12,13,14,15)]
d_m = d_m[,-1]
d_m = d_m[order(d_m$地块类型,d_m$种植季次,d_m$作物编号)]
d_m$`作物经济效益(万元)` = d_m$`产量(斤)` * d_m$`平均单价(元/斤)` - d_m$`种植成本/(元/亩)` * d_m$`种植面积/亩`
d_m$`利润率(%)` = d_m$`作物经济效益(万元)` / d_m$`成本(元)` * 100
write_xlsx(d_m,path="成本和产量.xlsx")


# 附件2补全
f2_all = d_total[,-9]
split_ranges <- strsplit(as.character(f2_all$`销售单价/(元/斤)`), "-")
f2_all$`单价下限(元/斤)` <- as.numeric(sapply(split_ranges, `[`, 1))
f2_all$`单价上限(元/斤)` <- as.numeric(sapply(split_ranges, `[`, 2))
f2_all$`平均单价(元/斤)` <- (f2_all$`单价下限(元/斤)`+f2_all$`单价上限(元/斤)`)/2
f2_all$序号 = as.numeric(f2_all$序号)
f2_all$作物编号 = as.numeric(f2_all$作物编号)
setkey(f2_all,作物名称)
f2_all = f2_all[name_var]
f2_all = f2_all[,c(2,3,12,4,5,6,7,8,9,10,11)]
f2_all = f2_all[order(f2_all$地块类型,f2_all$种植季次,f2_all$作物编号)]


avg_s = data.table(aggregate((`地块面积/亩`)~(地块类型),data=a,FUN = mean))
colnames(avg_s)[2] <- '平均地块面积/亩'
setkey(avg_s,地块类型)
setkey(f2_all,地块类型)
f2_all = f2_all[avg_s]
f2_all$`亩产经济效益/(元/亩)` = f2_all$`亩产量/斤` * f2_all$`平均单价(元/斤)` - f2_all$`种植成本/(元/亩)`
f2_all$`作物经济效益/(万元)` = f2_all$`亩产经济效益/(元/亩)` * f2_all$`平均地块面积/亩` / 10000
write_xlsx(f2_all,path="完整附件2_经济效益.xlsx")
