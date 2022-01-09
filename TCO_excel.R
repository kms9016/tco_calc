pacman::p_load(tidyverse, lubridate, reshape)

setwd("C:/Users/jason/OneDrive - Kakao Corp/1. KEP 업무/6. 아키텍처TCO")
getwd()
wb <- loadWorkbook("TCO분석_20211222.xlsx", create = TRUE)
############ 1. Ref_Flavor 작업 ###############
# Ref_Flavor 표 불러오기
#Ref_Flavor <- read.table("clipboard", header=TRUE)

Ref_Flavor <- readWorksheet(wb, 
                            sheet = "1. 비용항목정리", 
                            startRow = 4, endRow = 14,
                            startCol = 7, endCol = 13, 
                            header = TRUE)


# Ref_Flavor 구조확인
str(Ref_Flavor)

# Dataframe --> Tibble로 구조변경
Ref_Flavor <- as_tibble(Ref_Flavor)

# Ref_Flavor를 upivot
#melt(Ref_Flavor, id=c('Flavor'), na.rm=TRUE) %>% arrange(Flavor)
#gather(Ref_Flavor,attribute,value,'C3':'C2_320_Mlag') %>% arrange(Flavor)
Ref_Flavor <- pivot_longer(Ref_Flavor, cols = 'C3':'C2_320_Mlag', names_to='Type', values_to='VMs')


############ 2. Ref_구성타입 작업 #############
# Ref_구성타입 표 불러오기
#Ref_config <- read.table("clipboard", header=TRUE)
Ref_config <- readWorksheet(wb, 
                            sheet = "1. 비용항목정리", 
                            startRow = 4, endRow = 14,
                            startCol = 1, endCol = 5, 
                            header = TRUE)
# Ref_구성타입 구조확인
str(Ref_config)

# Dataframe --> Tibble로 구조변경
Ref_config <- as_tibble(Ref_config)

# Ref_구성타입과 Ref_Flavor병합
Ref_config <- left_join(Ref_config, Ref_Flavor, by=c('서버타입'='Type'))

# Ref_구성타입 컬럼 재조정
Ref_config <- select(Ref_config, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs)


############ 3. Qty_Rack 작업 #############
# Qty_Rack 표 불러오기
#Qty_Rack <- read.table("clipboard", header=TRUE)
Qty_Rack <- readWorksheet(wb, 
                            sheet = "1_1 항목별수량", 
                            startRow = 4, endRow = 14,
                            startCol = 1, endCol = 4, 
                            header = TRUE)


# Qty_Rack 구조확인
str(Qty_Rack)

# Dataframe --> Tibble로 구조변경
Qty_Rack <- as_tibble(Qty_Rack)

# Qty_Rack와 Ref_구성타입 병합
Qty_Rack <- left_join(Qty_Rack, Ref_config, by=c('구성타입'))

# Qty_Rack 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "랙수량"
Qty_Rack <- select(Qty_Rack, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, 랙수량)



############ 4. Qty_Server 작업 #############
# Qty_Server 표 불러오기
#Qty_Server <- read.table("clipboard", header=TRUE)
Qty_Server <- readWorksheet(wb, 
                          sheet = "1_1 항목별수량", 
                          startRow = 4, endRow = 14,
                          startCol = 6, endCol = 9, 
                          header = TRUE)

# Qty_Server 구조확인
str(Qty_Server)

# Dataframe --> Tibble로 구조변경
Qty_Server <- as_tibble(Qty_Server)

# Qty_Server Ref_구성타입 병합
Qty_Server <- left_join(Qty_Server, Ref_config, by=c('구성타입'))

# Qty_Server 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "서버수량"
Qty_Server <- select(Qty_Server, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, 서버수량)


############ 5. Qty_PRD 작업 #############
# Qty_PRD 표 불러오기
#Qty_PRD <- read.table("clipboard", header=TRUE)
Qty_PRD <- readWorksheet(wb, 
                            sheet = "1_1 항목별수량", 
                            startRow = 4, endRow = 14,
                            startCol = 11, endCol = 14, 
                            header = TRUE)


# Qty_PRD 구조확인
str(Qty_PRD)

# Dataframe --> Tibble로 구조변경
Qty_PRD <- as_tibble(Qty_PRD)

# Qty_PRD Ref_구성타입 병합
Qty_PRD <- left_join(Qty_PRD, Ref_config, by=c('구성타입'))

# Qty_PRD 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "PRD수"
Qty_PRD <- select(Qty_PRD, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, PRD수)



############ 6. Qty_CLM 작업 #############
# Qty_CLM 표 불러오기
#Qty_CLM <- read.table("clipboard", header=TRUE)
Qty_CLM <- readWorksheet(wb, 
                         sheet = "1_1 항목별수량", 
                         startRow = 4, endRow = 14,
                         startCol = 16, endCol = 19, 
                         header = TRUE)


# Qty_CLM 구조확인
str(Qty_CLM)

# Dataframe --> Tibble로 구조변경
Qty_CLM <- as_tibble(Qty_CLM)

# Qty_CLM Ref_구성타입 병합
Qty_CLM <- left_join(Qty_CLM, Ref_config, by=c('구성타입'))

# Qty_CLM 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "CLM수"
Qty_CLM <- select(Qty_CLM, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, CLM수)



############ 7. Qty_STG 작업 #############
# Qty_STG 표 불러오기
#Qty_STG <- read.table("clipboard", header=TRUE)
Qty_STG <- readWorksheet(wb, 
                         sheet = "1_1 항목별수량", 
                         startRow = 4, endRow = 14,
                         startCol = 21, endCol = 24, 
                         header = TRUE)

# Qty_STG 구조확인
str(Qty_STG)

# Dataframe --> Tibble로 구조변경
Qty_STG <- as_tibble(Qty_STG)

# Qty_STG Ref_구성타입 병합
Qty_STG <- left_join(Qty_STG, Ref_config, by=c('구성타입'))

# Qty_STG 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "STG수"
Qty_STG <- select(Qty_STG, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, STG수)



############ 8. Qty_ENV 작업 #############
# Qty_ENV 표 불러오기
#Qty_ENV <- read.table("clipboard", header=TRUE)
Qty_ENV <- readWorksheet(wb, 
                         sheet = "1_1 항목별수량", 
                         startRow = 4, endRow = 14,
                         startCol = 26, endCol = 29, 
                         header = TRUE)

# Qty_ENV 구조확인
str(Qty_ENV)

# Dataframe --> Tibble로 구조변경
Qty_ENV <- as_tibble(Qty_ENV)

# Qty_ENV Ref_구성타입 병합
Qty_ENV <- left_join(Qty_ENV, Ref_config, by=c('구성타입'))

# Qty_ENV 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "GBIC수", "Fiber수", "UTP수"
Qty_ENV <- select(Qty_ENV, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, GBIC수, Fiber수, UTP수)



############ 9. Qty_Rack + Qty_Server + Qty_PRD + Qty_CLM + Qty_STG + Qty_ENV 병합 작업 #############

# Qty_Rack / Qty_Sever 병합 후 Data_Qty로 저장
Data_Qty <- left_join(Qty_Rack, Qty_Server, by=c('Flavor','구성타입','서버타입'))

# Data_Qty에서 필요한 컬럼만 선택
Data_Qty <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU.x, MEM.x, 사용랙.x, VMs.x, 랙수량, 서버수량)

# Data_Qty / Qty_PRD 병합 후 Data_Qty로 저장
Data_Qty <- left_join(Data_Qty, Qty_PRD, by=c('Flavor','구성타입','서버타입'))

# Data_Qty에서 필요한 컬럼만 선택
Data_Qty <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU.x, MEM.x, 사용랙.x, VMs.x, 랙수량, 서버수량, PRD수)

# Data_Qty / Qty_CLM 병합 후 Data_Qty로 저장
Data_Qty <- left_join(Data_Qty, Qty_CLM, by=c('Flavor','구성타입','서버타입'))

# Data_Qty에서 필요한 컬럼만 선택
Data_Qty <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU.x, MEM.x, 사용랙.x, VMs.x, 랙수량, 서버수량, PRD수, CLM수)

# Data_Qty / Qty_STG 병합 후 Data_Qty로 저장
Data_Qty <- left_join(Data_Qty, Qty_STG, by=c('Flavor','구성타입','서버타입'))

# Data_Qty에서 필요한 컬럼만 선택
Data_Qty <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU.x, MEM.x, 사용랙.x, VMs.x, 랙수량, 서버수량, PRD수, CLM수, STG수)

# Data_Qty / Qty_ENV 병합 후 Data_Qty로 저장
Data_Qty <- left_join(Data_Qty, Qty_ENV, by=c('Flavor','구성타입','서버타입'))

# Data_Qty에서 필요한 컬럼만 선택
Data_Qty <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU.x, MEM.x, 사용랙.x, VMs.x, 랙수량, 
                   서버수량, PRD수, CLM수, STG수, GBIC수, Fiber수, UTP수)

# 컬럼명 바꾸기
colnames(Data_Qty)[4] <- "CPU"
colnames(Data_Qty)[5] <- "MEM"
colnames(Data_Qty)[6] <- "사용랙"
colnames(Data_Qty)[7] <- "서버당VMs"

#Data_Qty <- mutate(Data_Qty, ID=paste(Data_Qty$Flavor, Data_Qty$구성타입, Data_Qty$서버타입, sep=""))

############ 10. Price_Item 작업 #############

# Price_Item 표 불러오기 (숫자 데이터는 일반으로 속성 변경)
#Price_Item <- read.table("clipboard", header=TRUE)
Price_Item <- readWorksheet(wb, 
                         sheet = "2. 단가정리", 
                         startRow = 3, endRow = 15,
                         startCol = 1, endCol = 12, 
                         header = TRUE)

# Price_Item 구조확인
str(Price_Item)

# Dataframe --> Tibble로 구조변경
Price_Item <- as_tibble(Price_Item)

# ID컬럼 제외
#Price_Item <- select(Price_Item, -ID)

# Price_Item를 upivot
Price_Item <- pivot_longer(Price_Item, cols = '상면비':'UTP', names_to='비용구분', values_to='금액')


############ 10. Data_Qty_상면 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_상면 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 랙수량)

# 상면비용만 정리
Data_Qty_상면 <- Data_Qty_상면 %>% 
                  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
                  filter(비용구분 == "상면비")

colnames(Data_Qty_상면)[8] <- "수량"

############ 11. Data_Qty_전기 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_전기 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 랙수량)

# 전기비용만 정리
Data_Qty_전기 <- Data_Qty_전기 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "전기료")

colnames(Data_Qty_전기)[8] <- "수량"

############ 12. Data_Qty_랙구매 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_랙구매 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 랙수량)

# 랙구매비용만 정리
Data_Qty_랙구매 <- Data_Qty_랙구매 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "랙구매")

colnames(Data_Qty_랙구매)[8] <- "수량"

############ 13. Data_Qty_서버 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_서버 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 서버수량)

# 서버비용만 정리
Data_Qty_서버 <- Data_Qty_서버 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "서버")

colnames(Data_Qty_서버)[8] <- "수량"

############ 14. Data_Qty_PRD #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_PRD <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, PRD수)

# PRD비용만 정리
Data_Qty_PRD <- Data_Qty_PRD %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "PRD")

colnames(Data_Qty_PRD)[8] <- "수량"

############ 15. Data_Qty_CLM #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_CLM <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, CLM수)

# CLM비용만 정리
Data_Qty_CLM <- Data_Qty_CLM %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "CLM")

colnames(Data_Qty_CLM)[8] <- "수량"

############ 16. Data_Qty_STG #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_STG <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, STG수)

# STG비용만 정리
Data_Qty_STG <- Data_Qty_STG %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "STG")

colnames(Data_Qty_STG)[8] <- "수량"

############ 17. Data_Qty_GBIC #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_GBIC <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, GBIC수)

# GBIC비용만 정리
Data_Qty_GBIC <- Data_Qty_GBIC %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "GBIC")

colnames(Data_Qty_GBIC)[8] <- "수량"

############ 18. Data_Qty_FIBER #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_FIBER <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, Fiber수)

# Fiber비용만 정리
Data_Qty_FIBER <- Data_Qty_FIBER %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "Fiber")

colnames(Data_Qty_FIBER)[8] <- "수량"

############ 19. Data_Qty_UTP #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_UTP <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, UTP수)

# UTP비용만 정리
Data_Qty_UTP <- Data_Qty_UTP %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "UTP")

colnames(Data_Qty_UTP)[8] <- "수량"


############ 20. Data_Baseline 작업 #############

Data_Baseline <- bind_rows(Data_Qty_상면, Data_Qty_전기, Data_Qty_랙구매, Data_Qty_서버, 
                Data_Qty_PRD, Data_Qty_CLM, Data_Qty_STG, Data_Qty_GBIC, Data_Qty_FIBER, Data_Qty_UTP)

View(Data_Baseline)

#write.csv(Data_Baseline, file="Data_Baseline.csv")



############ 21. Flavor_1_set 작업 ###############
# Flavor_1_set 표 불러오기
#Flavor_1_set <- read.table("clipboard", header=TRUE)
Flavor_1_set <- readWorksheet(wb, 
                            sheet = "1_2 Flavor정리", 
                            startRow = 4, endRow = 14,
                            startCol = 1, endCol = 5, 
                            header = TRUE)

# Flavor_1_set 구조확인
str(Flavor_1_set)

# Dataframe --> Tibble로 구조변경
Flavor_1_set <- as_tibble(Flavor_1_set)


############ 22. Flavor_2_set 작업 ###############
# Flavor_2_set 표 불러오기
#Flavor_2_set <- read.table("clipboard", header=TRUE)
Flavor_2_set <- readWorksheet(wb, 
                              sheet = "1_2 Flavor정리", 
                              startRow = 4, endRow = 14,
                              startCol = 7, endCol = 11, 
                              header = TRUE)

# Flavor_2_set 구조확인
str(Flavor_2_set)

# Dataframe --> Tibble로 구조변경
Flavor_2_set <- as_tibble(Flavor_2_set)


############ 23. Flavor_3_set 작업 ###############
# Flavor_3_set 표 불러오기
#Flavor_3_set <- read.table("clipboard", header=TRUE)
Flavor_3_set <- readWorksheet(wb, 
                              sheet = "1_2 Flavor정리", 
                              startRow = 4, endRow = 14,
                              startCol = 13, endCol = 17, 
                              header = TRUE)

# Flavor_3_set 구조확인
str(Flavor_3_set)

# Dataframe --> Tibble로 구조변경
Flavor_3_set <- as_tibble(Flavor_3_set)


############ 24. Flavor_4_set 작업 ###############
# Flavor_4_set 표 불러오기
#Flavor_4_set <- read.table("clipboard", header=TRUE)
Flavor_4_set <- readWorksheet(wb, 
                              sheet = "1_2 Flavor정리", 
                              startRow = 4, endRow = 14,
                              startCol = 19, endCol = 23, 
                              header = TRUE)

# Flavor_4_set 구조확인
str(Flavor_4_set)

# Dataframe --> Tibble로 구조변경
Flavor_4_set <- as_tibble(Flavor_4_set)


############ 25. Flavor_5_set 작업 ###############
# Flavor_5_set 표 불러오기
#Flavor_5_set <- read.table("clipboard", header=TRUE)
Flavor_5_set <- readWorksheet(wb, 
                              sheet = "1_2 Flavor정리", 
                              startRow = 4, endRow = 14,
                              startCol = 25, endCol = 29, 
                              header = TRUE)

# Flavor_5_set 구조확인
str(Flavor_5_set)

# Dataframe --> Tibble로 구조변경
Flavor_5_set <- as_tibble(Flavor_5_set)


############ 26. Flavor_6_set 작업 ###############
# Flavor_6_set 표 불러오기
#Flavor_6_set <- read.table("clipboard", header=TRUE)
Flavor_6_set <- readWorksheet(wb, 
                              sheet = "1_2 Flavor정리", 
                              startRow = 4, endRow = 14,
                              startCol = 31, endCol = 35, 
                              header = TRUE)

# Flavor_6_set 구조확인
str(Flavor_6_set)

# Dataframe --> Tibble로 구조변경
Flavor_6_set <- as_tibble(Flavor_6_set)


############ 27. Flavor_1_set ~ Flavor_6_set 행결합 ###############
Data_Flavor <- bind_rows(Flavor_1_set, Flavor_2_set, Flavor_3_set, Flavor_4_set, Flavor_5_set, Flavor_6_set)
colnames(Data_Flavor)[2] <- "구성타입"


############ 28. 사용자 함수 정의 ###############

#비용항목 컬럼추가를 위한 사용자함수 생성
Category<-function(x){
  ifelse(x %in% c("상면비","전기료","랙구매"),"상면비용",
         ifelse(x %in% c("서버"),"서버",
                ifelse(x %in% c("PRD","CLM","STG"),"네트워크비용",
                       ifelse(x %in% c("GBIC","Fiber","UTP"),"환경구성비용","기타"))))
}

############ 29. Data_Baseline + Data_Flavor 병합 ###############
#Data_Baseline + Data_Flavor를 조인
#비용항목 컬럼 추가
#컬럼순서 조정
Data_Baseline <- Data_Baseline %>% 
                  left_join(Data_Flavor, by=c('Flavor','구성타입')) %>% 
                  mutate(비용항목 = Category(비용구분)) %>% 
                  select(비용항목, 비용구분, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 세트수량, 수량, 금액)


############ 30. Data_Baseline 전처리 작업 완료 ###############
#합계 컬럼 생성 (세트수량 * 항목별수량 * 단가 * 60개월)
Data_Baseline<-Data_Baseline %>%
  mutate(합계 = 세트수량 * 수량 * 금액 * 60)

# 총계 확인
Data_Baseline %>% 
  summarise(합계=sum(합계,na.rm=T), .groups='drop')

# 비용항목별 총계
Data_Baseline %>% 
  group_by(비용항목) %>% 
  summarise(합계=sum(합계,na.rm=T), .groups='drop')

# 비용구분별 총계
Data_Baseline %>% 
  group_by(비용항목, 비용구분) %>% 
  summarise(합계=sum(합계,na.rm=T), .groups='drop')

#write.csv(Data_Baseline, file="Data_Baseline.csv")


############ 31. Data_Baseline + Qty_Server 병합 ###############

#서버수량 참조테이블 생성
Ref_Server <- Qty_Server %>% 
              group_by(구성타입, 서버타입) %>% 
              summarise(서버수량=서버수량) %>% 
              unique()

#Data_Baseline과 Ref_Server를 조인하셔 서버수량 컬럼 추가
Data_Baseline <- Data_Baseline %>% 
  left_join(Ref_Server, by=c('구성타입','서버타입')) %>% 
  select(비용항목, 비용구분, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버수량, 서버당VMs, 세트수량, 수량, 금액)


write.csv(Data_Baseline, file="Data_Baseline.csv", row.names = FALSE)
            