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
 # a$`Number of abortions` = case_when(a$`Number of abortions` == 0 ~ "No",
 #                                     a$`Number of abortions` == 1 ~ 'Once',
 #                                     a$`Number of abortions`==2 ~ 'Twice or more',
 #                                     a$`Number of abortions` == 3 ~ 'Twice or more',
 #                                     a$`Number of abortions`== 5 ~ 'Twice or more')
 # table(a$`Number of abortions`)
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

# p <- a$PSQI1
# 
# write.csv(p,file = "D:/浙一/心理睡眠质量文章/产科科研/a.csv")




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
table(a$BMI)
factorVars <- c("Symptom","Temperature","Education","Occupation","COVID-19","Previous infection","Basic disease",
  "Hospital","febrifuge","COVID-19 vaccine","Flu vaccine",'Bad habits',"Abortion",'anxiety',"PSQI",
  'History of cesarean section','Miscarriage',"Gestational hypertension","Gestational diabetes",
  "Pregnancy with thrombocytopenia","embolism","Antiphospholipid syndrome","hypoproteinemia","Anemia","Iron deficiency",
 'First pregnancy','Number of full-term births','Number of premature births','Number of abortions','Number of children')
c <- colnames(a)
vars <- c(c[2],c[5:19],c[195],c[82],c[87:91],c[94:101],c[178:186],c[203],c[197],c[199:201],c[58:59],c[80:81])
#tableOne <- CreateTableOne(vars = vars, strata = "Covid-19 infection", data = a, factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, data = a, factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne
tableOne2 <- CreateTableOne(vars = vars, data = a,strata ='Miscarriage', factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne2
x <- print(tableOne)
x2 <- print(tableOne2)





#PQSI和zung焦虑分组
library(tableone)
library(readxl)
library(dplyr)
c <- colnames(a)
factorVars <- c("Symptom","Temperature","Education","Occupation","COVID-19","Previous infection","Basic disease",
                "Hospital","febrifuge","COVID-19 vaccine","Flu vaccine",'Bad habits',"Abortion",'anxiety',"PQSI",
                'History of cesarean section','Threatened abortion',"Gestational hypertension","Gestational diabetes",
                "Pregnancy with thrombocytopenia","embolism","Antiphospholipid syndrome","hypoproteinemia","Anemia","Iron deficiency",
                'First pregnancy','Number of full-term births','Number of premature births','Number of abortions','Number of children')

vars <- c(c[179],c[181:183],c[185:186],c[188],c[189],c[58:59],c[80:81])
#tableOne <- CreateTableOne(vars = vars, strata = "Covid-19 infection", data = a, factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, data = a, factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne
x <- print(tableOne)
tableOne2 <- CreateTableOne(vars = vars, data = a, factorVars = factorVars,strata ='Threatened abortion') #strata?൱??ָ??y??��??
tableOne2
x2 <- print(tableOne2)


#有没有新冠和有没有先兆流产的焦虑，睡眠质量，四个图的panel

library(dplyr)
library(ggpubr)
library(ggsci)

#小提琴图，等数据全了换几个变量试试
a$SAS <- a$Zung标准分
a$SDS <- a$SDS标准分
a$PSQI <- a$PSQI1
a$`COVID-19 infection` <- a$`COVID-19`
table(a$`COVID-19 infection`)
colnames(a)
a$`Number of abortions2`

a$`COVID-19 infection` <- case_when(a$`COVID-19 infection` == 1 ~ "YES",
                       a$`COVID-19 infection`==2 ~ "NO",
                       a$`COVID-19 infection`==3 ~ "unclear")
p1 <- ggplot(a, aes(Miscarriage, SAS)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")


p2 <- ggplot(a, aes(Miscarriage, SDS)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")
a$PSQI1
p3 <- ggplot(a, aes(Miscarriage, PSQI)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")

p4 <- ggplot(a, aes(Miscarriage, BMI)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")


p5 <- ggplot(a, aes(Miscarriage, Age)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")
a$`Number of abortions`
a$`History of miscarriage` <- a$`Number of abortions`
p6 <- ggplot(a, aes(Miscarriage, `History of miscarriage`)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")


ggarrange(p1,p2,p3,p4,p5,p6,ncol=3,nrow = 2, labels = c("A","B","C","D","E","F"))
colnames(a)
ggsave('~/Figure 1.pdf',height =6, width =10, dpi = 35)

#这里可以用箱式图分组，分不同年龄，不同新冠感染情况等
library(ggplot2)
table(a$Miscarriage,a$Education)
table(a$Miscarriage,a$`COVID-19`)
table(a$Miscarriage,a$Occupation)
table(a$Miscarriage,a$`First pregnancy`)
table(a$Miscarriage,a$`History of cesarean section`)
table(a$Miscarriage,a$`Basic disease`)
factorVars <- c("Education","Occupation","COVID-19","Basic disease",
                'History of cesarean section',
                'First pregnancy')
c <- colnames(a)
vars <- c(c[5:7],c[12],c[18:19])
#tableOne <- CreateTableOne(vars = vars, strata = "Covid-19 infection", data = a, factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, data = a, factorVars = factorVars,strata = "Miscarriage" ) #strata?൱??ָ??y??��??
tableOne

pl <- ggplot (a1,aes(Education,`Prevalence of miscarriage (%)`,fil1 =Education))+
  geom_boxplot(aes(Education,`Prevalence of miscarriage (%)`))+theme_classic()+scale_fill_lancet()+
  theme(legend.position = c(.8,.8))



ggarrange(p1,p2)



#回归模型+森林图
#先天流产的影响因素,有流产史的会影响先天流产
#crude model,只有先兆流产和睡眠，抑郁和焦虑的关系,森林图有三个，分别对应三个model的
#然后每个森林图里参考lancet那篇有圆圈和方框，一个是多因素的一个是单因素的，再画一个RCS图
#PSQI,SAS,SDS均是定量资料
a$Miscarriage <- case_when(a$Miscarriage == "Yes" ~ 1,
                                    a$Miscarriage=="No" ~ 0)
model11 <- glm(`Miscarriage`~Zung标准分, family="binomial", data=a)
model12 <- glm(`Miscarriage`~PSQI1, family="binomial", data=a)
model13 <- glm(`Miscarriage`~SDS标准分, family="binomial", data=a)

summary(model11)
summary(model12)
summary(model13)
c11 <- exp(cbind("Odds ratio" = coef(model11), confint.default(model11, level = 0.95)))
c12 <- exp(cbind("Odds ratio" = coef(model12), confint.default(model12, level = 0.95)))
c13 <- exp(cbind("Odds ratio" = coef(model13), confint.default(model13, level = 0.95)))

table(a$`COVID-19`)
table(a$`Basic disease`)
colnames(a)
a$`D-dimer`

model2 <- glm(`Miscarriage`~Age+factor(Occupation)+factor(Education)+`NEU%`+PLT+BMI+
                factor(`Basic disease`)+factor(`First pregnancy`)+factor(`history of cesarean section`)+factor(operation)+`Zung standard score`+PSQI1+`SDSstandard score`, family="binomial", data=a)


summary(model2)
c2 <- exp(cbind("Odds ratio" = coef(model2), confint.default(model2, level = 0.95)))


#列线图
library(rms)
library(rms)
be<-data.frame(a$Age,a$Occupation,a$Education,a$`Basic disease`,a$`First pregnancy`,a$PSQI1,a$SDS标准分,a$Zung标准分,a$Miscarriage)
dd <- datadist(be) #为后续程序设定数据环境
options(datadist='dd') #为后续程序设定数据环境
colnames(be)
be$a.Occupation <- as.factor(be$a.Occupation)
be$a.Education <- as.factor(be$a.Education)
be$a..Basic.disease. <- as.factor(be$a..Basic.disease.)

model2 <- lrm(a.Miscarriage~a.SDS标准分+a.Age+a.Occupation+a.Education+a..Basic.disease.+a..First.pregnancy.+a.history.of.cesarean.section+
                a.operation+a.BMI+a.NEU+a.PLT+a.Zung标准分+a.PSQI1, data=be)
model2 <- lrm(a.Miscarriage~a.SDS标准分+a.Age+
                a.Zung标准分+a.PSQI1, data=be)
summary(model2)

nom <- nomogram(model2, fun=plogis,
                fun.at=c(.01,.05,seq(.1,.9,by=.1),.95,.99,.999),
                maxscale=100,
                funlabel="risk")
plot(nom)

#ROC
a$Occupation <- as.factor(a$Occupation)
a$Education <- as.factor(a$Education)
a$`Basic disease` <- as.factor(a$`Basic disease`)
a$SAS <- a$Zung标准分
a$SDS <- a$SDS标准分
a$PSQI <- a$PSQI1
model2 <- glm(Miscarriage~Age+
                SAS+PSQI+SDS, family="binomial", data=a)
summary(model2)

predvalue<-predict(model2,type = "response",newdata = a)
b <- cbind(a$Miscarriage,predvalue)

library(pROC)
ROC <- roc(a$Miscarriage,predvalue)
ci(auc(ROC))
write.csv(as.matrix(coords(ROC, x="best", ret="all", transpose = FALSE)),
          "os.roc.train.csv")
pdf("roc.pdf",8,8)
plot(ROC,print.auc=TRUE,auc.polygon=F,lwd=3, max.auc.polygon=F, print.thres=F)
dev.off()

library(pROC)

plot.roc(a$Miscarriage,predvalue, # data
         
         percent=TRUE, # show all values in percent
         
         partial.auc=c(100, 90), partial.auc.correct=TRUE, # define a partial AUC (pAUC)
         
         print.auc=TRUE, #display pAUC value on the plot with following options:
         
         print.auc.pattern="Corrected pAUC (100-90%% SP):\n%.1f%%", print.auc.col="#87CEFA",
         
         auc.polygon=TRUE, auc.polygon.col="#87CEFA", # show pAUC as a polygon
         
         max.auc.polygon=TRUE, max.auc.polygon.col="#1c61b622", # also show the 100% polygon
         
         main="Partial AUC (pAUC)")

plot.roc(a$Miscarriage,predvalue,
  
         
         percent=TRUE, add=TRUE, type="n", # add to plot, but don't re-add the ROC itself (useless)
         
         partial.auc=c(100, 90), partial.auc.correct=TRUE,
         
         partial.auc.focus="se", # focus pAUC on the sensitivity
         
         print.auc=TRUE, print.auc.pattern="Corrected pAUC (100-90%% SE):\n%.1f%%", print.auc.col="#FFB6C1",
         
         print.auc.y=40, # do not print auc over the previous one
         
         auc.polygon=TRUE, auc.polygon.col="#FFB6C1",
         
         max.auc.polygon=TRUE, max.auc.polygon.col="#00860022")


rocobj <- plot.roc(a$Miscarriage,predvalue,
                   
                   main="Confidence intervals", percent=TRUE,
                   
                   ci=TRUE, # compute AUC (of AUC by default)
                   
                   print.auc=TRUE) # print the AUC (will contain the CI)

ciobj <- ci.se(rocobj, # CI of sensitivity
               
               specificities=seq(0, 100, 5)) # over a select set of specificities

plot(ciobj, type="shape", col="#FFB6C1") # plot as a blue shape

plot(ci(rocobj, of="thresholds", thresholds="best")) # add one threshold
#校准曲线
library(rms)
a$Occupation <- as.factor(a$Occupation)
a$Education <- as.factor(a$Education)
a$`Basic disease` <- as.factor(a$`Basic disease`)
a$SAS <- a$Zung标准分
a$SDS <- a$SDS标准分
a$PSQI <- a$PSQI1
model2 <- lrm(Miscarriage~Age+
                SAS+PSQI+SDS,data=a,x=T,y=T)

cal<-calibrate(model2,method = "boot",B=200)
?calibrate
plot(cal,
     xlim = c(0,1),
     xlab = "Predicted Probability",
     ylab="Observed  Probability",
     legend =FALSE,
     subtitles = FALSE)
abline(0,1,col="#00A087CC",lty=2,lwd=2)
lines(cal[,c("predy","calibrated.orig")],type="l",lwd=2,col="#FF7256",pch=16)
lines(cal[,c("predy","calibrated.corrected")],type="l",lwd=2,col="#4DBBD5CC",pch=16)
legend(0.55,0.35,
       c("Ideal","Apparent","Bias-corrected"),
       lty = c(2,1,1),
       lwd = c(2,2,2),
       col = c("#00A087CC","#FF7256","#4DBBD5CC"),
       bty="n") #"o"为加边框
#有颜色的列线图
library(VRPM)
a$Occupation <- as.factor(a$Occupation)
a$Education <- as.factor(a$Education)
a$`Basic disease` <- as.factor(a$`Basic disease`)
a$SAS <- a$Zung标准分
a$SDS <- a$SDS标准分
a$PSQI <- a$PSQI1
model2 <- glm(Miscarriage~Age+
              SAS+PSQI+SDS, family="binomial", data=a)
summary(model2)
colplot(model2,#构建的模型
        coloroptions=3,# 颜色方案，多种可选(1-5)
        risklabel ="Risk",# 概率标轴的名称
        filename="Nomogram") # 图片的文件名



#森林图
library(forestplot)
library(forestploter)
library(readxl)
od <- read_excel('~/forest.xlsx')# forest.csv

forestplot(labeltext = as.matrix(od[,1:2]),
           mean = od$est,
           lower = od$low,
           upper = od$hi,
           is.summary = c(F,F,F,F,F,F),#T,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F
           #T,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F
           zero = 1,
           boxsize = 0.1,
           colgap = unit(5,'mm'),
           lineheight = unit(1,'mm'),
           lwd.zero = 1.5,
           lwd.ci = 1.5,
           hrzl_lines=T,
           col = fpColors(box=c('#458B00','#458B00','#97FFFF','#97FFFF','black','black'),  lines = 'black', zero = 'black'),
           xlab = 'The estimates',lwd.xaxis = 2,
           graph.pos = 2)

w

#中介分析和中介图

# 中介效应介绍的  https://gabriellajg.github.io/EPSY-579-R-Cookbook-for-SEM/lavaan-lab-2-mediation-and-indirect-effects.html

#利用lavaan包跑中介效用模型
library(lavaan)
b <- as.data.frame(matrix(nrow=665))
b$SAS <- a$SAS
b$PSQI <- a$PSQI
b$Miscarriage <- a$Miscarriage
b$SDS <- a$SDS
b$Miscarriage
#anxiety是中介
labData <- b
ex1MediationSyntax <- "              #opening a quote
    #Regressions
    SAS ~ PSQI                    #M ~ X regression (a1 path)
    SDS ~ PSQI                    #M ~ X regression (a2 path)
    Miscarriage ~ PSQI + SAS+SDS         #Y ~ X + M regression (c prime and b1,b2)
    "  
ex1fit_freeX <- lavaan::sem(model = ex1MediationSyntax, data = labData, fixed.x = FALSE)
summary(ex1fit_freeX)
ex2MediationSyntax <- "                         #opening a quote
    #Regressions
      SAS ~ a1*PSQI                      #Label the a coefficient in the M regression.
      SDS~ a2*PSQI 
    Miscarriage ~ cPrime*PSQI + b1*SAS+b2*SDS   #Label the direct effect (cPrime) of X and direct effect of M (b) in the Y regression.
    " 
ex2fit <- lavaan::sem(model = ex2MediationSyntax, data = labData, fixed.x=FALSE)
summary(ex2fit)
ex3MediationSyntax <- "                         #opening a quote
    #Regressions
    SAS ~ a1*PSQI                      #Label the a coefficient in the M regression.
    SDS~ a2*PSQI 
    Miscarriage ~ cPrime*PSQI + b1*SAS+b2*SDS   #Label the direct effect (cPrime) of X and direct effect of M (b) in the Y regression.
    
    #Define New Parameters
    a1b1 := a1*b1                                   #the product term is computed as a*b
    a2b2:= a2*b2   
    c := cPrime + a1b1+a2b2                        #having defined ab, we can use this here.
" 
ex3fit <- lavaan::sem(model = ex3MediationSyntax, data = labData, fixed.x=FALSE)
summary(ex3fit)
summary(ex3fit, standardized = TRUE) #This includes standardized estimates. std.all contains usual regression standardization.
summary(ex3fit, ci = T)  #Include confidence intervals
summary(ex3fit, standardized = TRUE, ci = T)
set.seed(2022)
ex3Boot <- lavaan::sem(model = ex3MediationSyntax, data = labData, se = "bootstrap", bootstrap = 1000, fixed.x=FALSE) 
summary(ex3Boot, ci = TRUE) 
parameterEstimates(ex3Boot, level = 0.95, boot.ci.type="bca.simple")
library(semPlot)

semPaths(ex3Boot, what='est', 
         rotation = 2, # default rotation = 1 with four options
         curve = 2, # pull covariances' curves out a little
         nCharNodes = 0,
         nCharEdges = 0, # don't limit variable name lengths
         sizeMan = 8, # font size of manifest variable names
         style = "lisrel", # single-headed arrows vs. # "ram"'s 2-headed for variances
         edge.label.cex=1.2, curvePivot = TRUE, 
         fade=FALSE)


#SDS是中介
labData <- b
ex1MediationSyntax <- "              #opening a quote
    #Regressions
    SDS ~ PSQI                    #M ~ X regression (a path)
    Miscarriage ~ PSQI + SDS          #Y ~ X + M regression (c prime and b)
    "  
ex1fit_freeX <- lavaan::sem(model = ex1MediationSyntax, data = labData, fixed.x = FALSE)
summary(ex1fit_freeX)
ex2MediationSyntax <- "                         #opening a quote
    #Regressions
      SDS ~ a*PSQI                      #Label the a coefficient in the M regression.
    Miscarriage ~ cPrime*PSQI + b*SDS   #Label the direct effect (cPrime) of X and direct effect of M (b) in the Y regression.
    " 
ex2fit <- lavaan::sem(model = ex2MediationSyntax, data = labData, fixed.x=FALSE)
summary(ex2fit)
ex3MediationSyntax <- "                         #opening a quote
    #Regressions
    SDS ~ a*PSQI                      #Label the a coefficient in the M regression.
    Miscarriage ~ cPrime*PSQI + b*SDS   #Label the direct effect (cPrime) of X and direct effect of M (b) in the Y regression.
    
    #Define New Parameters
    ab := a*b                                   #the product term is computed as a*b
    c := cPrime + ab                        #having defined ab, we can use this here.
" 
ex3fit <- lavaan::sem(model = ex3MediationSyntax, data = labData, fixed.x=FALSE)
summary(ex3fit)
summary(ex3fit, standardized = TRUE) #This includes standardized estimates. std.all contains usual regression standardization.
summary(ex3fit, ci = T)  #Include confidence intervals
summary(ex3fit, standardized = TRUE, ci = T)
set.seed(2022)
ex3Boot <- lavaan::sem(model = ex3MediationSyntax, data = labData, se = "bootstrap", bootstrap = 1000, fixed.x=FALSE) 
summary(ex3Boot, ci = TRUE) 
parameterEstimates(ex3Boot, level = 0.95, boot.ci.type="bca.simple")
library(semPlot)

semPaths(ex3Boot, what='est', 
         rotation = 2, # default rotation = 1 with four options
         curve = 2, # pull covariances' curves out a little
         nCharNodes = 0,
         nCharEdges = 0, # don't limit variable name lengths
         sizeMan = 8, # font size of manifest variable names
         style = "lisrel", # single-headed arrows vs. # "ram"'s 2-headed for variances
         edge.label.cex=1.2, curvePivot = TRUE, 
         fade=FALSE)






#table 2
c <- colnames(a)
a$anxiety
factorVars <- c("anxiety")
vars <- c("anxiety",'factor1','factor2','factor3','factor4','factor5','factor6','factor7','PQSI1',"Zung总分","Zung标准分","PQSI")
tableOne <- CreateTableOne(vars = vars, strata = "Threatened abortion", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, strata = "COVID-19", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, strata = "History of cesarean section", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, strata = "Abortion", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??
tableOne <- CreateTableOne(vars = vars, strata = "First pregnancy", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??

tableOne
table(a$anxiety)
summary(tableOne)
print(tableOne,test = T)

#另外加表
table(a$age1)
a$age1  =  case_when(a$age<20 ~ "1",
                     a$age<35 ~ "2",
                     a$age<45 ~ "3",
                     a$age<60 ~ "4",
                     a$age>59 ~ "5")


colnames(a)

#两两比较
mtcars_aov <- aov(a$PQSI1~factor(a$`Covid-19 infection`))
summary(mtcars_aov)
TukeyHSD(mtcars_aov)

mtcars_aov <- aov(a$factor5~factor(a$`Covid-19 infection`))
summary(mtcars_aov)
TukeyHSD(mtcars_aov)

mtcars_aov <- aov(a$factor7~factor(a$`Covid-19 infection`))
summary(mtcars_aov)
TukeyHSD(mtcars_aov)

mtcars_aov <- aov(a$Zung总分~factor(a$`Covid-19 infection`))
summary(mtcars_aov)
TukeyHSD(mtcars_aov)

library(multcomp)
a$`Covid-19 infection` <- as.factor(a$`Covid-19 infection`)
x <- table(a$anxiety,a$`Covid-19 infection`)
mtcars_aov <- chisq.test(x)


rht <- glht(mtcars_aov, linfct=mcp(`Covid-19 infection` =="dunnett"),alternative="two.side" )
?glht


vars <- c("anxiety",'factor1','factor2','factor3','factor4','factor5','factor6','factor7','PQSI1',"Zung总分","Zung标准分","Sleep")
tableOne <- CreateTableOne(vars = vars, data = a, factorVars = factorVars) #strata?൱??ָ??y??��??
tableOne

?tableone
#figure 1
library(ggplot2)
a$sex <- as.factor(a$sex)
a$`Covid-19 infection` <- as.factor(a$`Covid-19 infection`)
a$PQSI1 <- as.numeric(a$PQSI1)
a$anxiety<- as.factor(a$anxiety)

p1 <- ggplot(a, aes(x=anxiety,y=PQSI1,fill=`Covid-19 infection`)) +
  geom_boxplot(position=position_dodge(1))+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank())+
  theme(legend.position="top")
# 设置图例的名称, 重新定义新的标签名称
p1 <-p1 + scale_fill_discrete(name="COVID-19 infection",
                              breaks=c("1", "2", "3"),
                              labels=c("Yes", "No", "Unclear"))+
  labs(y = "Global PQSI scores")



a$workyear <- as.factor(a$workyear)
p2 <- ggplot(a, aes(x=anxiety,y=PQSI1,fill=workyear)) +
  geom_boxplot(position=position_dodge(1))+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank())+
  theme(legend.position="top")
p2 <-p2 + scale_fill_discrete(name="Wroking years",
                              breaks=c("1", "2", "3"),
                              labels=c("< 5", "5~10", "≥10"))+
  labs(y = "Global PQSI scores")



a$education <- as.factor(a$education)
p3 <- ggplot(a, aes(x=anxiety,y=PQSI1,fill=sex)) +
  geom_boxplot(position=position_dodge(1))+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank())+
  theme(legend.position="top")


a$`Covid-19 vaccine` <- as.factor(a$`Covid-19 vaccine`)
p4 <- ggplot(a, aes(x=anxiety,y=PQSI1,fill=`Covid-19 vaccine`)) +
  geom_boxplot(position=position_dodge(1))+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank())+
  theme(legend.position="top")
p4 <-p4 + scale_fill_discrete(name="COVID-19 vaccine",
                              breaks=c("1", "2", "3","4","5","6"),
                              labels=c("Unvaccinated", "One dose", "Two doses","Three doses", "Four doses", "Other"))+
  labs(y = "Global PQSI scores")




p5 <- ggplot(a, aes(x=anxiety,y=PQSI1,fill=sex)) +
  geom_boxplot(position=position_dodge(1))+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank())+
  theme(legend.position="top")
p5 <-p5 + scale_fill_discrete(name="Sex",
                              breaks=c("1", "2"),
                              labels=c("Male", "Female"))+
  labs(y = "Global PQSI scores")




p6 <- ggplot(a, aes(x=anxiety,y=PQSI1,fill=age1)) +
  geom_boxplot(position=position_dodge(1))+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank())+
  theme(legend.position="top")


library(ggpubr)

ggarrange(p1,p2,p5,p4,ncol = 2, nrow = 2,
          labels = c("A","B","C","D"))
ggsave('~/Figure 1.pdf',height =6, width =10, dpi = 35)

head(a)

#睡眠几个维度
co <- c('factor1','f','factor3','factor4','f1','factor6','f2','PQSI1','Zung标准分')

a1 <- a[,co]
colnames(a1) <- c('factor1','factor2','factor3','factor4','factor5','factor6','factor7','PQSI','Zung')

cormat <- round(cor(a1),2)
head(cormat)
library(reshape2)
melted_cormat <- melt(cormat)
head(melted_cormat)
library(ggplot2)
ggplot(data = melted_cormat, aes(x=Var1, y=Var2, fill=value)) + 
  geom_tile()



# Get lower triangle of the correlation matrix
get_lower_tri<-function(cormat){
  cormat[upper.tri(cormat)] <- NA
  return(cormat)
}
# Get upper triangle of the correlation matrix
get_upper_tri <- function(cormat){
  cormat[lower.tri(cormat)]<- NA
  return(cormat)
}
upper_tri <- get_upper_tri(cormat)
upper_tri
# Melt the correlation matrix
library(reshape2)
melted_cormat <- melt(upper_tri, na.rm = TRUE)
# Heatmap
library(ggplot2)
ggplot(data = melted_cormat, aes(Var2, Var1, fill = value))+
  geom_tile(color = "white")+
  scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
                       midpoint = 0, limit = c(-1,1), space = "Lab", 
                       name="Pearson\nCorrelation") +
  theme_minimal()+ 
  theme(axis.text.x = element_text(angle = 45, vjust = 1, 
                                   size = 12, hjust = 1))+
  coord_fixed()



reorder_cormat <- function(cormat){
  # Use correlation between variables as distance
  dd <- as.dist((1-cormat)/2)
  hc <- hclust(dd)
  cormat <-cormat[hc$order, hc$order]
}
# Reorder the correlation matrix
cormat <- reorder_cormat(cormat)
upper_tri <- get_upper_tri(cormat)
# Melt the correlation matrix
melted_cormat <- melt(upper_tri, na.rm = TRUE)
# Create a ggheatmap
ggheatmap <- ggplot(melted_cormat, aes(Var2, Var1, fill = value))+
  geom_tile(color = "white")+
  scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
                       midpoint = 0, limit = c(-1,1), space = "Lab", 
                       name="Pearson\nCorrelation") +
  theme_minimal()+ # minimal theme
  theme(axis.text.x = element_text(angle = 45, vjust = 1, 
                                   size = 12, hjust = 1))+
  coord_fixed()
# Print the heatmap
print(ggheatmap)
ggheatmap + 
  geom_text(aes(Var2, Var1, label = value), color = "black", size = 4) +
  theme(
    axis.title.x = element_blank(),
    axis.title.y = element_blank(),
    panel.grid.major = element_blank(),
    panel.border = element_blank(),
    panel.background = element_blank(),
    axis.ticks = element_blank(),
    legend.justification = c(1, 0),
    legend.position = c(0.6, 0.7),
    legend.direction = "horizontal")+
  guides(fill = guide_colorbar(barwidth = 7, barheight = 1,
                               title.position = "top", title.hjust = 0.5))


#焦虑、睡眠质量差的影响因素分析
#quality为1是代表睡眠质量差
table(a$quality)
a$quality  =  case_when(a$PQSI1> 7 ~ "1",
                        a$PQSI1<8 ~ "0")

#table 3 单因素
a$age1  =  case_when(a$age<20 ~ "1",
                     a$age<35 ~ "2",
                     a$age<45 ~ "3",
                     a$age<60 ~ "4",
                     a$age>59 ~ "5")
a$education= case_when (a$education=="3"~"1",
                        a$education=="4"~"1",
                        a$education=="5"~"2")
table(a$education)
factorVars <- c("anxiety","education","Covid-19 infection","Covid-19 vaccine","workyear","sex","basic disease","flu infection","flu vaccine","age1")
vars <- c('PQSI1',"education","Covid-19 infection","Covid-19 vaccine","workyear","age1","sex","basic disease","flu infection","flu vaccine")
tableOne1 <- CreateTableOne(vars = vars, strata = "anxiety", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??
tableOne1
x <- print(tableOne1)


factorVars <- c("anxiety","education","Covid-19 infection","Covid-19 vaccine","workyear","sex","basic disease","flu infection","flu vaccine","age1")
vars <- c('anxiety',"education","Covid-19 infection","Covid-19 vaccine","workyear","sex","basic disease","flu infection","flu vaccine","age1")
tableOne2 <- CreateTableOne(vars = vars, strata = "Sleep", data = a, factorVars = factorVars,test=TRUE) #strata?൱??ָ??y??��??
tableOne2
x2 <- print(tableOne2)



#table 4多因素
#age分组，<20,20~35,35~45,45~60,大于60
table(a$age1)
a$age1  =  case_when(a$age<20 ~ "1",
                     a$age<35 ~ "2",
                     a$age<45 ~ "3",
                     a$age<60 ~ "4",
                     a$age>59 ~ "5")

a$`Covid-19 vaccine1`=case_when(a$`Covid-19 vaccine`==1 ~ "1",
                                a$`Covid-19 vaccine`==2 ~ "2",
                                a$`Covid-19 vaccine`==3 ~ "2",
                                a$`Covid-19 vaccine`==4 ~ "2",
                                a$`Covid-19 vaccine`==5 ~ "2",
                                a$`Covid-19 vaccine`==6 ~ "2")
table(a$`Covid-19 vaccine1`)
a$quality <- as.factor(a$quality)
model1 <- glm(quality~factor(sex)+factor(age1)+factor(workyear)+factor(education)+factor(`Covid-19 infection`)+
                factor(`basic disease`)+factor(`hospital`)+factor(`flu infection`)+
                factor(`Covid-19 vaccine`)+factor(`flu vaccine`),
              family="binomial", data=a)
x <- print(summary(model1) $ coefficients)


model2 <- glm(anxiety~factor(sex)+factor(age1)+factor(workyear)+factor(education)+factor(`Covid-19 infection`)+
                factor(`basic disease`)+factor(`hospital`)+factor(`flu infection`)+
                factor(`Covid-19 vaccine`)+factor(`flu vaccine`),
              family="binomial", data=a)
summary(model2)
x2 <- print(summary(model2) $ coefficients)




#炎症指标和流产的关系图
#WBC,NEU%,NEU,PLT,ALB,D-dimer,CRP,gestational_age
a$`D-dimer`
p11 <- ggplot(a, aes(Miscarriage, WBC)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")


p12 <- ggplot(a, aes(Miscarriage, NEU)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")

p13 <- ggplot(a, aes(Miscarriage, PLT)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")


p14 <- ggplot(a, aes(Miscarriage, ALB)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")

p15 <- ggplot(a, aes(Miscarriage, `D-dimer`)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")

p16 <- ggplot(a, aes(Miscarriage, CRP)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")



p17 <- ggplot(a, aes(Miscarriage, `NEU%`)) + geom_violin(aes(fill =Miscarriage))+
  geom_jitter(height = 0, width = 0.1)+stat_compare_means()+
  scale_color_lancet()+
  theme_bw()+theme(panel.grid.major = element_blank(),
                   panel.grid.minor.x =element_blank(),
                   legend.position = "none")

ggarrange(p11,p12,p13,p14,p15,p16,p17,ncol = 4, nrow = 2,
          labels = c("A","B","C","D","E","F","G"))
ggsave('~/Figure S3.pdf',height =6, width =10, dpi = 35)




#cox
library(dplyr)
library(rms)
library(survival)
library(pec)

options(datadist='ddist')
options(scipen = 999)
a$PoorSleep
a$Age <- as.numeric(a$Age)

a$Miscarriage <- case_when(a$Miscarriage == "Yes" ~ 1,
                           a$Miscarriage=="No" ~ 0)

cox1 <- coxph(Surv(gestational_age,Miscarriage)~Age+factor(Occupation)+factor(Education)+`NEU%`+PLT+BMI+
                factor(`Basic disease`)+factor(`First pregnancy`)+factor(`history of cesarean section`)+factor(operation)+`Zung standard score`+PSQI1+`SDS standard score`, data=a)
cox11<- summary(cox1)
cox11$coefficients

#竞争风险模型
library(tidycmprsk)
library(gtsummary)
library(ggsurvfit)
a <- a[1:444,]
#0是无，1是早产，2是流产
a$Status1 <- case_when(a$Status == 1 ~ 2,
                       a$Status==2 ~ 1,
                      a$Status==0 ~ 0)

a$Status <- as.factor(a$Status)
cuminc(Surv(gestational_age, Status) ~ 1, data = a)
cuminc(Surv(gestational_age, Status) ~ 1, data = a) %>%
  ggcuminc() +
  labs(x = "Days") +
  add_confidence_interval() +
  add_risktable()
cuminc(Surv(gestational_age,Status) ~ 1, data = a) %>%
  ggcuminc(outcome = c("1", "2")) +
  ylim(c(0, 1)) +
  labs(x = "Days")
library(tidycmprsk)

cuminc(Surv(gestational_age,Status) ~  Age+Occupation+Education+`NEU%`+PLT+BMI+`history of cesarean section`+operation
         `Basic disease`+`First pregnancy`+`Zung standard score`+PSQI1+`SDS standard score`, data = a) %>%
  tbl_cuminc(
    times = 1826.25,
    label_header = "xx") %>%
  add_p()
cuminc(Surv(gestational_age,Status) ~ Age+Occupation+Education+`NEU%`+PLT+BMI+`history of cesarean section`+operation
       `Basic disease`+`First pregnancy`+`Zung standard score`+PSQI1+`SDS standard score`, data = a) %>%
  ggcuminc() +
  labs( x = "Weeks") +
  add_confidence_interval() +
  add_risktable()
comp_risk_model <- crr(Surv(gestational_age,Status) ~  Age+Occupation+Education+`NEU%`+PLT+BMI+`history of cesarean section`+operation
                       `Basic disease`+`First pregnancy`+`Zung standard score`+PSQI1+`SDS standard score`,data = a)
crr(Surv(gestational_age, Status) ~ Age+Occupation+Education+`NEU%`+PLT+BMI+`history of cesarean section`+operation
    `Basic disease`+`First pregnancy`+`Zung standard score`+PSQI1+`SDS standard score`, data = a) %>%
  tbl_regression(exp = TRUE)

#敏感性分析
a<-read_excel("~/Datat.xlsx")
#前面数据清洗的步骤再运行一遍
a$Miscarriage <- case_when(a$Miscarriage == "Yes" ~ 1,
                           a$Miscarriage=="No" ~ 0)
a$Age <- as.numeric(a$Age)
model2 <- glm(`Miscarriage`~Age+factor(Occupation)+factor(Education)+`NEU%`+PLT+BMI+
                factor(`Basic disease`)+factor(`First pregnancy`)+factor(`history of cesarean section`)+factor(operation)+`Zung standard score`+PSQI1+`SDS standard score`, family="binomial", data=a)
summary(model2)
c2 <- exp(cbind("Odds ratio" = coef(model2), confint.default(model2, level = 0.95)))


#敏感性分析森林图
library(forestplot)
library(forestploter)
od <- read_excel('~/forest.xlsx')# forest.csv

forestplot(labeltext = as.matrix(od[,1:2]),
           mean = od$est,
           lower = od$low,
           upper = od$hi,
           is.summary = c(F,F,F,F,F,F),#T,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F
           #T,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F
           zero = 1,
           boxsize = 0.1,
           colgap = unit(5,'mm'),
           lineheight = unit(1,'mm'),
           lwd.zero = 1.5,
           lwd.ci = 1.5,
           hrzl_lines=T,
           col = fpColors(box=c('#458B00','#458B00','#97FFFF','#97FFFF','black','black'),  lines = 'black', zero = 'black'),
           xlab = 'The estimates',lwd.xaxis = 2,
           graph.pos = 2)

od <- read_excel('~/forest.xlsx',sheet = "Sheet2")# forest.csv

forestplot(labeltext = as.matrix(od[,1:2]),
           mean = od$est,
           lower = od$low,
           upper = od$hi,
           is.summary = c(F,F,F,F,F,F),#T,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F
           #T,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F
           zero = 1,
           boxsize = 0.1,
           colgap = unit(5,'mm'),
           lineheight = unit(1,'mm'),
           lwd.zero = 1.5,
           lwd.ci = 1.5,
           hrzl_lines=T,
           col = fpColors(box=c('#458B00','#458B00','#97FFFF','#97FFFF','black','black'),  lines = 'black', zero = 'black'),
           xlab = 'The estimates',lwd.xaxis = 2,
           graph.pos = 2)


od <- read_excel('~/forest.xlsx',sheet = "Sheet3")# forest.csv

forestplot(labeltext = as.matrix(od[,1:2]),
           mean = od$est,
           lower = od$low,
           upper = od$hi,
           is.summary = c(F,F,F,F,F,F),#T,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F
           #T,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F
           zero = 1,
           boxsize = 0.1,
           colgap = unit(5,'mm'),
           lineheight = unit(1,'mm'),
           lwd.zero = 1.5,
           lwd.ci = 1.5,
           hrzl_lines=T,
           col = fpColors(box=c('#458B00','#458B00','#97FFFF','#97FFFF','black','black'),  lines = 'black', zero = 'black'),
           xlab = 'The estimates',lwd.xaxis = 2,
           graph.pos = 2)



od <- read_excel('~/forest.xlsx',sheet = "Sheet4")# forest.csv

forestplot(labeltext = as.matrix(od[,1:2]),
           mean = od$est,
           lower = od$low,
           upper = od$hi,
           is.summary = c(F,F,F,F,F,F),#T,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F
           #T,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F,T,F,F,F
           zero = 1,
           boxsize = 0.1,
           colgap = unit(5,'mm'),
           lineheight = unit(1,'mm'),
           lwd.zero = 1.5,
           lwd.ci = 1.5,
           hrzl_lines=T,
           col = fpColors(box=c('#458B00','#458B00','#97FFFF','#97FFFF','black','black'),  lines = 'black', zero = 'black'),
           xlab = 'The estimates',lwd.xaxis = 2,
           graph.pos = 2)
