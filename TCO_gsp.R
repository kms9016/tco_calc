pacman::p_load(tidyverse, lubridate, reshape, XLConnect, googlesheets4)

setwd("C:/Users/kms90/OneDrive - Kakao Corp/1. KEP 업무/6. 아키텍처TCO")
getwd()

wb <- "https://docs.google.com/spreadsheets/d/1vaMukLB1djuXzSNJil8v-wCjr4_a_CzVxGqmwuf04zE/edit?usp=sharing"
gs4_auth(email = "jason.gang@kakaoenterprise.com")

############ 1. Ref_Flavor 작업 ###############
# Ref_Flavor 표 불러오기
Ref_Flavor <- range_read(wb, range="Ref_Flavor")

# Ref_Flavor 구조확인
str(Ref_Flavor)

# Dataframe --> Tibble로 구조변경
Ref_Flavor <- as_tibble(Ref_Flavor)

# Ref_Flavor를 upivot
Ref_Flavor <- pivot_longer(Ref_Flavor, cols = 'C3':'C2_320_Mlag', names_to='Type', values_to='VMs')


# Sheet_add로 Ref_Flavor 저장
range_delete(wb, range="Sheet_add!A2:C", shift = "up")
range_write(wb, data = Ref_Flavor, range="Sheet_add!A2")


############ 2. Ref_Config 작업 #############
# Ref_Config 표 불러오기
Ref_Config <- range_read(wb, range="Ref_Config")

# Ref_구성타입 구조확인
str(Ref_Config)

# Dataframe --> Tibble로 구조변경
Ref_Config <- as_tibble(Ref_Config)

# Ref_구성타입과 Ref_Flavor병합
Ref_Config <- left_join(Ref_Config, Ref_Flavor, by=c('서버타입'='Type'))

# Ref_구성타입 컬럼 재조정
Ref_Config <- select(Ref_Config, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs)

# Sheet_add로 Ref_Config 저장
range_delete(wb, range="Sheet_add!F2:L", shift = "up")
range_write(wb, data = Ref_Config, range="Sheet_add!F2")


############ 3. Qty_Rack 작업 #############
# Qty_Rack 표 불러오기
Qty_Rack <- range_read(wb, range="Qty_Rack")

# Qty_Rack 구조확인
str(Qty_Rack)

# Dataframe --> Tibble로 구조변경
Qty_Rack <- as_tibble(Qty_Rack)

# Qty_Rack와 Ref_구성타입 병합
Qty_Rack <- left_join(Qty_Rack, Ref_Config, by=c('구성타입'))

# Qty_Rack 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "랙수량"
Qty_Rack <- select(Qty_Rack, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, 랙수량)

# Sheet_add로 Qty_Rack 저장
range_delete(wb, range="Sheet_add!O2:V", shift = "up")
range_write(wb, data = Qty_Rack, range="Sheet_add!O2")


############ 4. Qty_Server 작업 #############
# Qty_Server 표 불러오기
Qty_Server <- range_read(wb, range="Qty_Server")

# Qty_Server 구조확인
str(Qty_Server)

# Dataframe --> Tibble로 구조변경
Qty_Server <- as_tibble(Qty_Server)

# Qty_Server Ref_구성타입 병합
Qty_Server <- left_join(Qty_Server, Ref_Config, by=c('구성타입'))

# Qty_Server 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "서버수량"
Qty_Server <- select(Qty_Server, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, 서버수량)

# Sheet_add로 Qty_Server 저장
range_delete(wb, range="Sheet_add!Y2:AF", shift = "up")
range_write(wb, data = Qty_Server, range="Sheet_add!Y2")


############ 5. Qty_PRD 작업 #############
# Qty_PRD 표 불러오기
Qty_PRD <- range_read(wb, range="Qty_PRD")

# Qty_PRD 구조확인
str(Qty_PRD)

# Dataframe --> Tibble로 구조변경
Qty_PRD <- as_tibble(Qty_PRD)

# Qty_PRD Ref_구성타입 병합
Qty_PRD <- left_join(Qty_PRD, Ref_Config, by=c('구성타입'))

# Qty_PRD 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "PRD수"
Qty_PRD <- select(Qty_PRD, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, PRD수)

# Sheet_add로 Qty_PRD 저장
range_delete(wb, range="Sheet_add!AI2:AP", shift = "up")
range_write(wb, data = Qty_PRD, range="Sheet_add!AI2")


############ 6. Qty_CLM 작업 #############
# Qty_CLM 표 불러오기
Qty_CLM <- range_read(wb, range="Qty_CLM")

# Qty_CLM 구조확인
str(Qty_CLM)

# Dataframe --> Tibble로 구조변경
Qty_CLM <- as_tibble(Qty_CLM)

# Qty_CLM Ref_구성타입 병합
Qty_CLM <- left_join(Qty_CLM, Ref_Config, by=c('구성타입'))

# Qty_CLM 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "CLM수"
Qty_CLM <- select(Qty_CLM, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, CLM수)

# Sheet_add로 Qty_CLM 저장
range_delete(wb, range="Sheet_add!AS2:AZ", shift = "up")
range_write(wb, data = Qty_CLM, range="Sheet_add!AS2")


############ 7. Qty_STG 작업 #############
# Qty_STG 표 불러오기
Qty_STG <- range_read(wb, range="Qty_STG")

# Qty_STG 구조확인
str(Qty_STG)

# Dataframe --> Tibble로 구조변경
Qty_STG <- as_tibble(Qty_STG)

# Qty_STG Ref_구성타입 병합
Qty_STG <- left_join(Qty_STG, Ref_Config, by=c('구성타입'))

# Qty_STG 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "STG수"
Qty_STG <- select(Qty_STG, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, STG수)

# Sheet_add로 Qty_STG 저장
range_delete(wb, range="Sheet_add!AS2:AZ", shift = "up")
range_write(wb, data = Qty_STG, range="Sheet_add!BC2")


############ 8. Qty_ENV 작업 #############
# Qty_ENV 표 불러오기
Qty_ENV <- range_read(wb, range="Qty_ENV")

# Qty_ENV 구조확인
str(Qty_ENV)

# Dataframe --> Tibble로 구조변경
Qty_ENV <- as_tibble(Qty_ENV)

# Qty_ENV Ref_구성타입 병합
Qty_ENV <- left_join(Qty_ENV, Ref_Config, by=c('구성타입'))

# Qty_ENV 컬럼 재조정
#"Flavor", "구성타입", "서버타입", "CPU", "MEM", "사용랙", "서버VMs", "GBIC수", "Fiber수", "UTP수"
Qty_ENV <- select(Qty_ENV, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, VMs, GBIC수, Fiber수, UTP수)

# Sheet_add로 Qty_ENV 저장
range_delete(wb, range="Sheet_add!BM2:BV", shift = "up")
range_write(wb, data = Qty_ENV, range="Sheet_add!BM2")


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

# Sheet_add로 Data_Qty 저장
range_delete(wb, range="Sheet_add!BY2:CM", shift = "up")
range_write(wb, data = Data_Qty, range="Sheet_add!BY2")


############ 10. Price_Item 작업 #############

# Price_Item 표 불러오기 (숫자 데이터는 일반으로 속성 변경)
Price_Item <- range_read(wb, range="Price_Item")

# Price_Item 구조확인
str(Price_Item)

# Dataframe --> Tibble로 구조변경
Price_Item <- as_tibble(Price_Item)

# Price_Item를 upivot
Price_Item <- pivot_longer(Price_Item, cols = '상면비':'UTP', names_to='비용구분', values_to='금액')

# Sheet_add로 Price_Item를 저장
range_delete(wb, range="Sheet_add!CP2:CS", shift = "up")
range_write(wb, data = Price_Item, range="Sheet_add!CP2")


############ 11. Data_Qty_상면 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_상면 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 랙수량)

# 상면비용만 정리
Data_Qty_상면 <- Data_Qty_상면 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "상면비")

colnames(Data_Qty_상면)[8] <- "수량"

# Sheet_add로 Data_Qty_상면 저장
range_delete(wb, range="Sheet_add!CV2:DE", shift = "up")
range_write(wb, data = Data_Qty_상면, range="Sheet_add!CV2")


############ 12. Data_Qty_전기 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_전기 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 랙수량)

# 전기비용만 정리
Data_Qty_전기 <- Data_Qty_전기 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "전기료")

colnames(Data_Qty_전기)[8] <- "수량"

# Sheet_add로 Data_Qty_상면 저장
range_delete(wb, range="Sheet_add!DH2:DQ", shift = "up")
range_write(wb, data = Data_Qty_전기, range="Sheet_add!DH2")


############ 13. Data_Qty_랙구매 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_랙구매 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 랙수량)

# 랙구매비용만 정리
Data_Qty_랙구매 <- Data_Qty_랙구매 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "랙구매")

colnames(Data_Qty_랙구매)[8] <- "수량"

# Sheet_add로 Data_Qty_랙구매 저장
range_delete(wb, range="Sheet_add!DT2:EC", shift = "up")
range_write(wb, data = Data_Qty_랙구매, range="Sheet_add!DT2")

############ 14. Data_Qty_서버 #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_서버 <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, 서버수량)

# 서버비용만 정리
Data_Qty_서버 <- Data_Qty_서버 %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "서버")

colnames(Data_Qty_서버)[8] <- "수량"

# Sheet_add로 Data_Qty_서버 저장
range_delete(wb, range="Sheet_add!EF2:EO", shift = "up")
range_write(wb, data = Data_Qty_서버, range="Sheet_add!EF2")

############ 15. Data_Qty_PRD #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_PRD <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, PRD수)

# PRD비용만 정리
Data_Qty_PRD <- Data_Qty_PRD %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "PRD")

colnames(Data_Qty_PRD)[8] <- "수량"

# Sheet_add로 Data_Qty_PRD 저장
range_delete(wb, range="Sheet_add!ER2:FA", shift = "up")
range_write(wb, data = Data_Qty_PRD, range="Sheet_add!ER2")

############ 16. Data_Qty_CLM #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_CLM <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, CLM수)

# CLM비용만 정리
Data_Qty_CLM <- Data_Qty_CLM %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "CLM")

colnames(Data_Qty_CLM)[8] <- "수량"

# Sheet_add로 Data_Qty_CLM 저장
range_delete(wb, range="Sheet_add!FD2:FM", shift = "up")
range_write(wb, data = Data_Qty_CLM, range="Sheet_add!FD2")

############ 17. Data_Qty_STG #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_STG <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, STG수)

# STG비용만 정리
Data_Qty_STG <- Data_Qty_STG %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "STG")

colnames(Data_Qty_STG)[8] <- "수량"

# Sheet_add로 Data_Qty_STG 저장
range_delete(wb, range="Sheet_add!FO2:FX", shift = "up")
range_write(wb, data = Data_Qty_STG, range="Sheet_add!FO2")

############ 18. Data_Qty_GBIC #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_GBIC <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, GBIC수)

# GBIC비용만 정리
Data_Qty_GBIC <- Data_Qty_GBIC %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "GBIC")

colnames(Data_Qty_GBIC)[8] <- "수량"

# Sheet_add로 Data_Qty_GBIC 저장
range_delete(wb, range="Sheet_add!GA2:GJ", shift = "up")
range_write(wb, data = Data_Qty_GBIC, range="Sheet_add!GA2")

############ 19. Data_Qty_FIBER #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_FIBER <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, Fiber수)

# Fiber비용만 정리
Data_Qty_FIBER <- Data_Qty_FIBER %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "Fiber")

colnames(Data_Qty_FIBER)[8] <- "수량"

# Sheet_add로 Data_Qty_FIBER 저장
range_delete(wb, range="Sheet_add!GM2:GV", shift = "up")
range_write(wb, data = Data_Qty_FIBER, range="Sheet_add!GM2")

############ 20. Data_Qty_UTP #############

# Data_Qty에서 랙수량 컬럼만 가져옴
Data_Qty_UTP <- select(Data_Qty, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버당VMs, UTP수)

# UTP비용만 정리
Data_Qty_UTP <- Data_Qty_UTP %>% 
  left_join(Price_Item, by=c('서버타입','사용랙')) %>% 
  filter(비용구분 == "UTP")

colnames(Data_Qty_UTP)[8] <- "수량"

# Sheet_add로 Data_Qty_UTP 저장
range_delete(wb, range="Sheet_add!GY2:HH", shift = "up")
range_write(wb, data = Data_Qty_UTP, range="Sheet_add!GY2")


############ 21. Data_Baseline 작업 #############

Data_Baseline <- bind_rows(Data_Qty_상면, Data_Qty_전기, Data_Qty_랙구매, Data_Qty_서버, 
                           Data_Qty_PRD, Data_Qty_CLM, Data_Qty_STG, Data_Qty_GBIC, Data_Qty_FIBER, Data_Qty_UTP)


# Sheet_add로 Data_Baseline 저장
range_delete(wb, range="Sheet_add!HK2:HT", shift = "up")
range_write(wb, data = Data_Baseline, range="Sheet_add!HK2")


############ 22. Flavor_1_set 작업 ###############
# Flavor_1_set 표 불러오기
Flavor_1_set <- range_read(wb, range="Flavor_1_set")

# Flavor_1_set 구조확인
str(Flavor_1_set)

# Dataframe --> Tibble로 구조변경
Flavor_1_set <- as_tibble(Flavor_1_set)


############ 23. Flavor_2_set 작업 ###############
# Flavor_2_set 표 불러오기
Flavor_2_set <- range_read(wb, range="Flavor_2_set")

# Flavor_2_set 구조확인
str(Flavor_2_set)

# Dataframe --> Tibble로 구조변경
Flavor_2_set <- as_tibble(Flavor_2_set)


############ 24. Flavor_3_set 작업 ###############
# Flavor_3_set 표 불러오기
Flavor_3_set <- range_read(wb, range="Flavor_3_set")

# Flavor_3_set 구조확인
str(Flavor_3_set)

# Dataframe --> Tibble로 구조변경
Flavor_3_set <- as_tibble(Flavor_3_set)


############ 24. Flavor_4_set 작업 ###############
# Flavor_4_set 표 불러오기
Flavor_4_set <- range_read(wb, range="Flavor_4_set")

# Flavor_4_set 구조확인
str(Flavor_4_set)

# Dataframe --> Tibble로 구조변경
Flavor_4_set <- as_tibble(Flavor_4_set)


############ 25. Flavor_5_set 작업 ###############
# Flavor_5_set 표 불러오기
Flavor_5_set <- range_read(wb, range="Flavor_5_set")

# Flavor_5_set 구조확인
str(Flavor_5_set)

# Dataframe --> Tibble로 구조변경
Flavor_5_set <- as_tibble(Flavor_5_set)


############ 26. Flavor_6_set 작업 ###############
# Flavor_6_set 표 불러오기
Flavor_6_set <- range_read(wb, range="Flavor_6_set")

# Flavor_6_set 구조확인
str(Flavor_6_set)

# Dataframe --> Tibble로 구조변경
Flavor_6_set <- as_tibble(Flavor_6_set)


############ 27. Flavor_1_set ~ Flavor_6_set 행결합 ###############
Data_Flavor <- bind_rows(Flavor_1_set, Flavor_2_set, Flavor_3_set, Flavor_4_set, Flavor_5_set, Flavor_6_set)
colnames(Data_Flavor)[2] <- "구성타입"

# Sheet_add로 Data_Flavor 저장
range_delete(wb, range="Sheet_add!HW2:IA", shift = "up")
range_write(wb, data = Data_Flavor, range="Sheet_add!HW2")


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

# Sheet_add로 Data_Baseline 저장
range_delete(wb, range="Sheet_add!ID2:IO", shift = "up")
range_write(wb, data = Data_Baseline, range="Sheet_add!ID2")


############ 31. Data_Baseline + Qty_Server 병합 ###############

#서버수량 참조테이블 생성
Ref_Server <- Qty_Server %>% 
  group_by(구성타입, 서버타입) %>% 
  summarise(서버수량=서버수량) %>% 
  unique()

# Sheet_add로 Ref_Server 저장
range_delete(wb, range="Sheet_add!IS2:IU", shift = "up")
range_write(wb, data = Ref_Server, range="Sheet_add!IS2")

#Data_Baseline과 Ref_Server를 조인하셔 서버수량 컬럼 추가
Data_Baseline <- Data_Baseline %>% 
  left_join(Ref_Server, by=c('구성타입','서버타입')) %>% 
  select(비용항목, 비용구분, Flavor, 구성타입, 서버타입, CPU, MEM, 사용랙, 서버수량, 서버당VMs, 세트수량, 수량, 금액)


# Sheet_add로 Data_Baseline 저장
range_delete(wb, range="Sheet_add!IX2:JJ", shift = "up")
range_write(wb, data = Data_Baseline, range="Sheet_add!IX2")

range_delete(wb, range="5_Data_Baseline!A1:M", shift = "up")
range_write(wb, data = Data_Baseline, range="5_Data_Baseline!A1")
