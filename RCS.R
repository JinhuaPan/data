library(tidyverse)

library(Hmisc)

library(rms)

## 从《R语言实战（第二版）》上找了个可以做logistic回归的数据集

#install.packages("AER")

library(rlang)
library(lavaan)
library(semPlot)
library(readxl)
library("eeptools")
library(ggplot2)
library(dplyr)
#table 1
library(tableone)
a<-read_excel("~/Data.xlsx")
#职业分成在职和无职业，employed和unemployed,8是unemployed，其他合并成employed
a$Occupation = case_when(a$Occupation == 8 ~ 'unemployed',
                         a$Occupation == 1 ~ 'employed',
                         a$Occupation==2 ~ 'employed',
                         a$Occupation == 3 ~ 'employed',
                         a$Occupation == 4 ~ 'employed',
                         a$Occupation==5 ~ 'employed',
                         a$Occupation == 6 ~ 'employed',
                         a$Occupation == 7 ~ 'employed',
                         a$Occupation==9 ~ 'employed')
table(a$Occupation)
#教育程度1、2、3合并成高中及以下，大学本科及专科、硕士及以上
a$Education = case_when(a$ Education == 1 ~ 'Senior high school and below',
                        a$ Education == 2 ~ 'Senior high school and below',
                        a$Education==3 ~ 'Senior high school and below',
                        a$Education == 4 ~" Bachelor's degree",
                        a$Education== 5 ~ 'Master degree or above')
table(a$Education)
# 退烧药 febrifuge分成了吃和没吃两种
a$febrifuge = case_when(a$febrifuge == 1 ~ "No",
                        a$febrifuge == 2 ~ 'Yes',
                        a$febrifuge==3 ~ 'Yes',
                        a$febrifuge == 6 ~ 'Yes',
                        a$febrifuge== 7 ~ 'Yes')
table(a$febrifuge)
#这里如果需要重新跑表1的画要再跑一下
a$`Number of abortions` = case_when(a$`Number of abortions` == 0 ~ "No",
                                    a$`Number of abortions` == 1 ~ 'Once',
                                    a$`Number of abortions`==2 ~ 'Twice or more',
                                    a$`Number of abortions` == 3 ~ 'Twice or more',
                                    a$`Number of abortions`== 5 ~ 'Twice or more')
table(a$`Number of abortions`)
#有无吸烟饮酒等不良嗜好，吸烟和饮酒合并成yes，无是no
a$`Bad habits` = case_when(a$`Bad habits` == 1 ~ "No",
                           a$`Bad habits` == 2 ~ 'Yes',
                           a$`Bad habits`==3 ~ 'Yes')
table(a$`Bad habits`)

#有没有孩子，0是no，1是one or more
a$`Number of children` = case_when(a$`Number of children`  == 0 ~ "No",
                                   a$`Number of children`  == 1 ~ 'One or more',
                                   a$`Number of children` ==2 ~ 'One or more')
table(a$`Number of children` )

#疫苗接种情况，1是没有，2是接种两次及以下，3是接种三次以上
a$`COVID-19 vaccine` = case_when(a$`COVID-19 vaccine`  == 1 ~ "No",
                                 a$`COVID-19 vaccine`  == 2 ~ 'Twice or less',
                                 a$`COVID-19 vaccine`  ==3 ~ 'Twice or less',
                                 a$`COVID-19 vaccine`  == 4 ~ "Thrice or more",
                                 a$`COVID-19 vaccine`  == 5 ~ 'Thrice or more')
table(a$`COVID-19 vaccine` )

#第几次怀孕，1是第一次，2、3、4不是第一次
a$`First pregnancy` = case_when(a$`First pregnancy`  == 1 ~ "Yes",
                                a$`First pregnancy`   == 2 ~ 'No',
                                a$`First pregnancy`  ==3 ~ 'No',
                                a$`First pregnancy`  == 4 ~ "No")
table(a$`First pregnancy`)

#Number of abortions要不要改成是否有流产史，有和没有
a$`Number of abortions2` = case_when(a$`Number of abortions`  == 'No' ~ 'No',
                                     a$`Number of abortions`   == 'Once' ~ 'Yes',
                                     a$`Number of abortions`  =='Twice or more'~ 'Yes')
table(a$`Number of abortions2`)
a$factor1 <- a$Sleep6
a$factor1 = case_when(a$factor1 == 1 ~ 0,
                      a$factor1 ==2 ~ 1,
                      a$factor1 ==3 ~ 2,
                      a$factor1 ==4 ~ 3)

table(a$Sleep2)

a$Sleep22 = case_when(a$Sleep2 <3 ~ 0,
                      a$Sleep2 ==3 ~ 1,
                      a$Sleep2 ==4 ~ 2,
                      a$Sleep2 ==5 ~ 2,
                      a$Sleep2 >5 ~ 3)
table(a$Sleep22)
table(a$Sleep5A)
a$Sleep5A = case_when(a$Sleep5A == 1 ~ 0,
                      a$Sleep5A ==2 ~ 1,
                      a$Sleep5A ==3 ~ 2,
                      a$Sleep5A ==4 ~ 3)


a$factor2=a$Sleep5A+a$Sleep22
table(a$factor2)
a$factor2 <- case_when(a$factor2 == 0 ~ 0,
                       a$factor2 == 1  ~ 1,
                       a$factor2 ==2 ~ 1,
                       a$factor2 == 4 ~ 2,
                       a$factor2 == 3 ~ 2,
                       a$factor2 == 5 ~ 3,
                       a$factor2 == 6 ~ 3)
table(a$Sleep2)

a$factor3 <- case_when(a$Sleep4 > 7 ~ 0,
                       a$Sleep4 > 6  ~ 1,
                       a$Sleep4 > 5 ~ 2,
                       a$Sleep4 < 5.01 ~ 3)

table(a$factor3)
table(a$Sleep4)

a$factor4 <- a$Sleep22
table(a$Sleep5J)

a$Sleep5B = case_when(a$Sleep5B == 1 ~ 0,
                      a$Sleep5B ==2 ~ 1,
                      a$Sleep5B ==3 ~ 2,
                      a$Sleep5B ==4 ~ 3)
a$Sleep5C = case_when(a$Sleep5C == 1 ~ 0,
                      a$Sleep5C ==2 ~ 1,
                      a$Sleep5C ==3 ~ 2,
                      a$Sleep5C ==4 ~ 3)
a$Sleep5D = case_when(a$Sleep5D == 1 ~ 0,
                      a$Sleep5D ==2 ~ 1,
                      a$Sleep5D ==3 ~ 2,
                      a$Sleep5D ==4 ~ 3)
table(a$Sleep5E)
a$Sleep5E = case_when(a$Sleep5E == 1 ~ 0,
                      a$Sleep5E ==2 ~ 1,
                      a$Sleep5E ==3 ~ 2,
                      a$Sleep5E ==4 ~ 3,
                      a$Sleep5E ==5 ~ 3)
a$Sleep5F = case_when(a$Sleep5F == 1 ~ 0,
                      a$Sleep5F ==2 ~ 1,
                      a$Sleep5F ==3 ~ 2,
                      a$Sleep5F ==4 ~ 3)
a$Sleep5G = case_when(a$Sleep5G == 1 ~ 0,
                      a$Sleep5G ==2 ~ 1,
                      a$Sleep5G ==3 ~ 2,
                      a$Sleep5G ==4 ~ 3)
a$Sleep5H = case_when(a$Sleep5H == 1 ~ 0,
                      a$Sleep5H ==2 ~ 1,
                      a$Sleep5H ==3 ~ 2,
                      a$Sleep5H ==4 ~ 3)
a$Sleep5I = case_when(a$Sleep5I == 1 ~ 0,
                      a$Sleep5I ==2 ~ 1,
                      a$Sleep5I ==3 ~ 2,
                      a$Sleep5I ==4 ~ 3)
a$Sleep5J = case_when(a$Sleep5J == 1 ~ 0,
                      a$Sleep5J ==2 ~ 1,
                      a$Sleep5J ==3 ~ 2,
                      a$Sleep5J ==4 ~ 3,
                      a$Sleep5J ==5 ~ 3)
table(a$Sleep5J)

a$f1 <- a$Sleep5B+a$Sleep5C+a$Sleep5D+a$Sleep5E+a$Sleep5F+a$Sleep5G+a$Sleep5H+a$Sleep5I+a$Sleep5J
table(a$f1)

a$factor5 <- case_when(a$f1 ==0 ~ 0,
                       a$f1 <9.01 ~ 1,
                       a$f1 <18.01 ~ 2,
                       a$f1 <34  ~ 3)

table(a$factor5)


a$Sleep7 = case_when(a$Sleep7 == 1 ~ 0,
                     a$Sleep7 ==2 ~ 1,
                     a$Sleep7 ==3 ~ 2,
                     a$Sleep7 ==4 ~ 3)
a$factor6 <- a$Sleep7
a$Sleep8 = case_when(a$Sleep8 == 1 ~ 0,
                     a$Sleep8 ==2 ~ 1,
                     a$Sleep8 ==3 ~ 2,
                     a$Sleep8 ==4 ~ 3)
a$Sleep9 = case_when(a$Sleep9 == 1 ~ 0,
                     a$Sleep9 ==2 ~ 1,
                     a$Sleep9 ==3 ~ 2,
                     a$Sleep9 ==4 ~ 3)
a$f2 <- a$Sleep8+a$Sleep9
table(a$f2)

a$factor7 = case_when(a$f2 == 0 ~ 0,
                      a$f2 ==1  ~ 1,
                      a$f2 ==2  ~ 1,
                      a$f2 ==3 ~ 2,
                      a$f2 ==4  ~ 2,
                      a$f2 ==5  ~ 3,
                      a$f2 ==6  ~ 3)

table(a$factor7)

a$PSQI1 <- a$factor1+a$factor2+a$factor3+a$factor4+a$factor5+a$factor6+a$factor7
table(a$PSQI1)

a$PSQI=case_when(a$PSQI1< 4 ~ "很好",
                 a$PSQI1 <8.1 ~ "还行",
                 a$PSQI1 <16.1 ~ "一般",
                 a$PSQI1 <25 ~ "很差")

table(a$PSQI)
a$PoorSleep=case_when(a$PSQI1 < 5.1 ~ "No",
                      a$PSQI1> 5 ~ "Yes")
table(a$PoorSleep)


a$anxiety <- case_when(a$Zung标准分> 49 ~ "YES",
                       a$Zung标准分<50~ "NO")
table(a$anxiety)
a$depression <- case_when(a
                          $SDS标准分> 52 ~ "YES",
                          a$SDS标准分<53 ~ "NO")
table(a$depression)
table(a$`Miscarriage`)
#Mild, moderate, severe
table(a$anxiety)
a$Weight <- as.numeric(a$Weight)
a$Age <- as.numeric(a$Age)

a$BMI=a$Weight/(a$Height**2)
a$BMI1 <- case_when(a$BMI< 18.5 ~ "Underweight",
                    a$BMI<25 ~ "Normal weight",
                    a$BMI<30 ~ "Overweight",
                    a$BMI<80~ "Obese")
table(a$BMI1)/664


## PQSI的RCS
be<-data.frame(a$Age,a$Occupation,a$Education,a$`Basic disease`,a$`First pregnancy`,a$PSQI1,a$SDS标准分,a$Zung标准分,a$`Miscarriage`)
dd <- datadist(be) #为后续程序设定数据环境
options(datadist='dd') #为后续程序设定数据环境
# colnames(be) <- c("Age","Occupation","Education","Basic disease" ,
#                   "First pregnancy","PQSI1","SDS标准分",
#                   "Zung标准分","Threatened abortion")


fit <- lrm(a.Miscarriage ~ rcs(a.PSQI1,3)+a.Zung标准分+a.SDS标准分+a.Age+
             a.Occupation+a.Education+a..Basic.disease., data=be)  ##拟合年龄样条化后的logistic回归

fit <- lrm(a.Miscarriage ~ rcs(a.PSQI1,3), data=be)  ##拟合年龄样条化后的logistic回归

print(fit)
an<-anova(fit)
ggplot(Predict(fit, a.PSQI1, ref.zero=TRUE, fun=exp),anova=an, pval=T)

## 通过以上代码，图形已经大体出来了。但参照似乎默认是选的中位数（不确定，看图猜的），尝试修改参照组并添加一些元素美化

dd$limits$a.PSQI1[2] <-12###设定标准
fit1=update(fit)###更新模型
OR<-Predict(fit1, a.PSQI1,fun=exp,ref.zero = TRUE) ##生成预测值
P1 <- ggplot()+geom_line(data=OR, aes(a.PSQI1,yhat),linetype=1,size=1,alpha = 0.9,colour="red")+
  geom_ribbon(data=OR, aes(a.PSQI1,ymin = lower, ymax = upper),alpha = 0.3,fill="red")+
  geom_hline(yintercept=1, linetype=2,size=1)+theme_classic()+ 
  labs(title = "RCS", x="PSQI", y="OR(95%CI)")###绘图

## SDS的RCS
be<-data.frame(a$Age,a$Occupation,a$Education,a$`Basic disease`,a$`First pregnancy`,a$PSQI1,a$SDS标准分,a$Zung标准分,a$`Miscarriage`)
dd <- datadist(be) #为后续程序设定数据环境
options(datadist='dd') #为后续程序设定数据环境

colnames(be)
fit <- lrm(a.Miscarriage ~ rcs(a.SDS标准分,3), data=be)  ##拟合年龄样条化后的logistic回归

print(fit)
an<-anova(fit)
ggplot(Predict(fit,a.SDS标准分, ref.zero=TRUE, fun=exp),anova=an, pval=T)

## 通过以上代码，图形已经大体出来了。但参照似乎默认是选的中位数（不确定，看图猜的），尝试修改参照组并添加一些元素美化

dd$limits$a.SDS标准分[2] <-70###设定标准
fit1=update(fit)###更新模型
OR<-Predict(fit1, a.SDS标准分,fun=exp,ref.zero = TRUE) ##生成预测值
P2 <-ggplot()+geom_line(data=OR, aes(a.SDS标准分,yhat),linetype=1,size=1,alpha = 0.9,colour="red")+
  geom_ribbon(data=OR, aes(a.SDS标准分,ymin = lower, ymax = upper),alpha = 0.3,fill="red")+
  geom_hline(yintercept=1, linetype=2,size=1)+theme_classic()+ 
  labs(title = "RCS", x="SDS", y="OR(95%CI)")###绘图

## SAS的RCS
be<-data.frame(a$Age,a$Occupation,a$Education,a$`Basic disease`,a$`First pregnancy`,a$PSQI1,a$SDS标准分,a$Zung标准分,a$`Miscarriage`)
dd <- datadist(be) #为后续程序设定数据环境
options(datadist='dd') #为后续程序设定数据环境

colnames(be)
fit <- lrm(a.Miscarriage ~a.Zung标准分, data=be)  ##拟合年龄样条化后的logistic回归


print(fit)
an<-anova(fit)
ggplot(Predict(fit,a.Zung标准分, ref.zero=TRUE, fun=exp),anova=an, pval=T)

## 通过以上代码，图形已经大体出来了。但参照似乎默认是选的中位数（不确定，看图猜的），尝试修改参照组并添加一些元素美化

dd$limits$a.Zung标准分[2] <-65###设定标准
fit1=update(fit)###更新模型
OR<-Predict(fit1,a.Zung标准分,fun=exp,ref.zero = TRUE) ##生成预测值
P3 <-ggplot()+geom_line(data=OR, aes(a.Zung标准分,yhat),linetype=1,size=1,alpha = 0.9,colour="red")+
  geom_ribbon(data=OR, aes(a.Zung标准分,ymin = lower, ymax = upper),alpha = 0.3,fill="red")+
  geom_hline(yintercept=1, linetype=2,size=1)+theme_classic()+ 
  labs(title = "RCS", x="SAS", y="OR(95%CI)")###绘图





## Age的RCS
be<-data.frame(a$Age,a$Occupation,a$Education,a$`Basic disease`,a$`First pregnancy`,a$PSQI1,a$SDS标准分,a$Zung标准分,a$`Miscarriage`)
dd <- datadist(be) #为后续程序设定数据环境
options(datadist='dd') #为后续程序设定数据环境

colnames(be)
fit <- lrm(a.Miscarriage ~a.Age, data=be)  ##拟合年龄样条化后的logistic回归


print(fit)
an<-anova(fit)
ggplot(Predict(fit,a.Age, ref.zero=TRUE, fun=exp),anova=an, pval=T)

## 通过以上代码，图形已经大体出来了。但参照似乎默认是选的中位数（不确定，看图猜的），尝试修改参照组并添加一些元素美化

dd$limits$a.Age[2] <-30###设定标准
fit1=update(fit)###更新模型
OR<-Predict(fit1,a.Age,fun=exp,ref.zero = TRUE) ##生成预测值
P4 <-ggplot()+geom_line(data=OR, aes(a.Age,yhat),linetype=1,size=1,alpha = 0.9,colour="red")+
  geom_ribbon(data=OR, aes(a.Age,ymin = lower, ymax = upper),alpha = 0.3,fill="red")+
  geom_hline(yintercept=1, linetype=2,size=1)+theme_classic()+ 
  labs(title = "RCS", x="Age", y="OR(95%CI)")###绘图


ggarrange(P1,P2,P3,P4,ncol=2,nrow=2,label.x = 0.5,label.y=0.97,font.label = list(size = 7))


ggsave("???Ϲ??????Է?????ͼ.png",dpi=350,width=12,height=8)

