### 19,20 주석 / 40,41 추가 / 66,71추가 / 79,80 주석 81 추가 / 78추가+ 77띄어쓰기하기/ 83 추가+띄어쓰기하기 / 107변경 / 137,138 변경  
### E_BESS_DA와 E_BESS_RT에너지 제약조건 삭제 및 E_BESS추가(144~158) /194~196 (식32) 변경/ 식(33,34,35) 변경 및 질문 / 식(43~51) 변경 및 식(45~51)질문 예비력 램프레이트 + 쓰는 이유 (한 번 더 질문)
### 식(58,59)추가 / 식 61~63 수정 / B_t,C_t,BESS1_RT,BESS2_RT 수정 / 엑셀에 나타내고, 그래프출력(파이썬을 이용하는게 best 아니면 엑셀으로라도)
from __future__ import print_function
from cmath import inf
from docplex.mp.model import Model
from docplex.util.environment import get_environment
import win32com.client as win32
import pandas as pd
import os

### 엑셀 불러오기
excel = win32.Dispatch("Excel.Application")
wb1 = excel.Workbooks.Open(os.getcwd()+"\\Data\\robust model_data.xlsx")
Price_DA = wb1.Sheets("Price_DA")              # Day-ahead price
Price_RS = wb1.Sheets("Price_RS")              # Reserve prices
Price_UR = wb1.Sheets("Price_UR")              # Up regulation prices
Price_DR = wb1.Sheets("Price_DR")              # Down regulation prices
#Expected_P_UR = wb1.Sheets("Expected_P_UR")    # Expected deployed power in up regulation services
#Expected_P_DR = wb1.Sheets("Expected_P_DR")    # Expected deployed power in down regulation services
Expected_P_RT_WPR = wb1.Sheets("Expected_P_RT_WPR")  # Expected wind power realization

### 파라미터 설정
time_dim = 24     # 시간 개수 (t)
min_dim = 12      # ex) 5분 x 12 = 1시간 (j)
del_S = 1/min_dim # Duration of intra-hourly interval ex) 5min = 1/12(h)
BESS_dim = 2      # BESS 개수 (s)
WPR_dim = 1       # 풍력발전기 개수 (w)
Marginal_cost_CH = 1    # Marginal cost of BES in charging modes
Marginal_cost_DCH = 1   # Marginal cost of BES in discharging modes
Marginal_cost_WPR = 3   # Marginal cost of WPR
Ramp_rate_WPR = 3       # Ramp-rate of WPR (MW)
Initial_BESS = 15   # Initial energy of BESS (MWh)
E_min_BESS = [0, 0]    # Minimum energy of BESS (MWh)
E_max_BESS = [30, 30]   # Maximum energy of BESS (MWh)
P_max_BESS = [5,3]    # Maximum power of BESS (MW)
P_min_BESS = [0,0]    # Minimum power of BESS (MW)
Ramp_rate_BESS = [5,3]  # Ramp-rate of BESS

Robust_percent = 0   # Robust percent (0~1)
contri_reg_percent = 0.2 # Up regulation 혹은 down regulation에 기여하는 비율 (0~1)

### 최적화 파트
def build_optimization_model(name='Robust_Optimization_Model'):
    mdl = Model(name=name)   # Model - Cplex에 입력할 Model 이름 입력 및 Model 생성
    mdl.parameters.mip.tolerances.mipgap = 0.0001;   # 최적화 계산 오차 설정

    time = [t for t in range(1,time_dim+1)]    # (t)의 one dimension
    time_min = [(t,j) for t in range(1,time_dim + 1) for j in range(1,min_dim+1)]   # (t,j)의 two dimension
    time_n_BESS = [(t,j,s) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1)]   # (t,j,s)의 three dimension
    time_n_WPR = [(t,j,w) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1)]     # (t,j,w)의 three dimension

    ### Continous Variable 지정 (연속 변수, 실수 변수)
    ## Day-ahead
    P_DA_S = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-S")   # Selling bids in the day-ahead market
    P_DA_B = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-B")   # Buying bids in the day-ahead market
    P_RS = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-RS")       # Reserve bids
    
    P_UR = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-UR")   # Deployed power in the up-regulation services
    P_DR = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-DR")   # Deployed power in the down-regulation services

    P_DA_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-CH")     # Day-ahead scheduling of BES in charging modes
    P_DA_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-DCH")   # Day-ahead scheduling of BES in discharging modes
    P_DA_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DA-WPR")    # Day-ahead scheduling of WPR

    P_UR_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-UR-CH")    # Deployed up regulation power of BES in charging mode
    P_UR_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-UR-DCH")   # Deployed up regulation power of BES in discharging mode
    P_UR_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-UR-WPR")    # Deployed up regulation power of WPR   

    P_DR_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DR-CH")      # Deployed down regulation power of BES in charging mode
    P_DR_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DR-DCH")   # Deployed down regulation power of BES in discharging mode
    P_DR_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DR-WPR")     # Deployed down regulation power of WPR  

    P_RS_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-CH")      # Reserve scheduling of BES in charging modes
    P_RS_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-DCH")    # Reserve scheduling of BES in discharging modes
    P_RS_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RS-WPR")     # Reserve scheduling of WPR
    ## Real-time
    P_SP_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-SP-WPR")             # Spilled power of WPR (difference between the realization of wind power and the scheduled power of WPR)
    # E_BESS_DA = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="E-BESS-DA")        # Energy level of BES in Day-ahead
    # E_BESS_RT = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="E-BESS-RT")        # Energy level of BES in Real-time
    E_BESS = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="E-BESS")                # Energy level of BES in Real-time
    P_RT_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RT-WPR")             # Realization of wind power in real-time
    AV_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="AV-WPR")                 # Auxiliary variables for linearization    
    ### Functions
    AV_RO = mdl.continuous_var_dict(time, lb=0, ub=inf, name="AV-RO")            # Auxiliary variable of RO / Income in real-time
    AV_RO_DA = mdl.continuous_var_dict(time, lb=0, ub=inf, name="AV-RO-DA")      # Income in day-ahead
    B_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="B-t")                # Income function of owner
    C_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="C-t")                # Cost function of owner
    BESS1_DA = mdl.continuous_var_dict(time, lb=0, ub=inf, name="BESS1-DA")      # Income of BESS#1 in day-ahead
    BESS2_DA = mdl.continuous_var_dict(time, lb=0, ub=inf, name="BESS2-DA")      # Income of BESS#2 in day-ahead
    WPR_DA = mdl.continuous_var_dict(time, lb=0, ub=inf, name="WPR-DA")          # Income of WPR in day-ahead
    BESS1_RT = mdl.continuous_var_dict(time, lb=0, ub=inf, name="BESS1-RT")      # Income of BESS#1 in real-time
    BESS2_RT = mdl.continuous_var_dict(time, lb=0, ub=inf, name="BESS2-RT")      # Income of BESS#2 in real-time
    WPR_RT = mdl.continuous_var_dict(time, lb=0, ub=inf, name="WPR-RT")          # Income of WPR in real-time   

    ### Binary Variable 지정 (이진 변수)
    D_Char = mdl.binary_var_dict(time_n_BESS, name="D-Char")      # Charging binary variables of BES (알파)
    D_Dchar = mdl.binary_var_dict(time_n_BESS, name="D-DChar")    # Discharging binary variables of BES (베타)
    D_WPR = mdl.binary_var_dict(time_n_WPR, name="D-WPR")         # Commitment status binary variable of WPR
    
    ### Objective function - 식(1) / 식(65)
    mdl.maximize(mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         + Price_RS.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         - mdl.sum((mdl.sum(Marginal_cost_DCH * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH * del_S * P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(Marginal_cost_WPR * del_S * P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         + AV_RO[t] for t in range(1,time_dim+1)))

    ### Robust Optizimation을 위한 변수 (BESS + WPR) - 식(65)               
    mdl.add_constraints(AV_RO[t] <= mdl.sum(Price_UR.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_UR_DCH[(t,j,s)] - P_UR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) 
                                            + Price_DR.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_DR_CH[(t,j,s)] - P_DR_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))                                      
                                            - mdl.sum(Marginal_cost_DCH * del_S * (P_UR_DCH[(t,j,s)] + P_DR_DCH[(t,j,s)])  + Marginal_cost_CH * del_S * (P_UR_CH[(t,j,s)] + P_DR_CH[(t,j,s)])  for s in range(1,BESS_dim+1)) - mdl.sum(Marginal_cost_WPR * del_S * P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) 
                                            for j in range(1,min_dim+1)) for t in range(1,time_dim+1))
      
    ### Equality constraints - 식(4) ~ 식(6) + 식(12) ~ 식(14)
    # Day-ahead bids 식(4)~식(6)
    mdl.add_constraints(P_DA_DCH[(t,j,s)] == P_DA_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(4)

    mdl.add_constraints(P_DA_CH[(t,j,s)] == P_DA_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(5)

    mdl.add_constraints(P_DA_WPR[(t,j,w)] == P_DA_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(6)

    # Reserve bids 식(12)~식(14)
    mdl.add_constraints(P_RS_CH[(t,j,s)] == P_RS_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(12)   

    mdl.add_constraints(P_RS_DCH[(t,j,s)] == P_RS_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(13)                                           

    mdl.add_constraints(P_RS_WPR[(t,j,w)] == P_RS_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(14)

    ### Constraints of day-ahead energy / reserve bids / real-time deployed power in the up and down regulation services - 식(7) ~ 식(11), 
    # 식(7)-(9)는 논문이 틀림
    mdl.add_constraints(P_DA_S[t] == del_S * mdl.sum(mdl.sum(P_DA_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) for j in range(1,min_dim+1)) for t in range(1,time_dim+1))  # 식(7)

    mdl.add_constraints(P_DA_B[t] == del_S * mdl.sum(P_DA_CH[(t,j,s)] for j in range(1,min_dim+1) for s in range(1,BESS_dim+1)) for t in range(1,time_dim+1))  # 식(8)
    
    mdl.add_constraints(P_RS[t] == del_S * mdl.sum(mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) for j in range(1,min_dim+1)) for t in range(1,time_dim+1))  # 식(9)
    
    # 식(10)-(11) 변형
    mdl.add_constraints(P_UR[(t,j)] == mdl.sum(mdl.sum(P_UR_DCH[(t,j,s)] + P_UR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1) for t in range(1,time_dim+1))
    mdl.add_constraints(P_DR[(t,j)] == mdl.sum(mdl.sum(P_DR_DCH[(t,j,s)] + P_DR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1) for t in range(1,time_dim+1))
    
    # 식(15) ~ 식(16)
    mdl.add_constraints(P_UR[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(15)
    mdl.add_constraints(P_DR[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(16)
    
    ### Constarints of stored energy of BES - 식(17) ~ 식(19)
    ## 식(17) t>=1, j>=2
    mdl.add_constraints(E_BESS[(t,j,s)] == E_BESS[(t,j-1,s)] + del_S * (P_DA_CH[(t,j,s)] - P_DA_DCH[(t,j,s)] + P_DR_CH[(t,j,s)] + P_UR_CH[(t,j,s)] - P_UR_DCH[(t,j,s)] - P_DR_DCH[(t,j,s)]) 
                       for t in range(1, time_dim+1) for j in range(2, min_dim+1) for s in range(1,BESS_dim+1))

    ## 식(19) t>=2, j=1
    mdl.add_constraints(E_BESS[(t,j,s)] == E_BESS[(t-1,min_dim,s)] 
                       for t in range(2, time_dim+1) for j in range(1, 2) for s in range(1,BESS_dim+1))
    ## 식(19) t=1, j=1
    mdl.add_constraints(E_BESS[(t,j,s)] == E_max_BESS[s-1]/2 
                       for t in range(1, 2) for j in range(1, 2) for s in range(1,BESS_dim+1))
          
    ##식(19) t=T, j=Nj      
    mdl.add_constraints(E_BESS[(t,j,s)] == E_max_BESS[s-1]/2 
                       for t in range(time_dim, time_dim+1) for j in range(min_dim, min_dim+1) for J in range(min_dim, min_dim+1) for s in range(1,BESS_dim+1))
                        
    ### Constarints of capacity - 식(20) ~ 식(38)
    # Power capacity of BES in day-ahead planning - 식(20) ~ 식(25)
    # Deployed power of BES in regulation service - 식(26) ~ 식(31)
    for t in range(1,time_dim+1):
        for j in range(1,min_dim+1):
            for s in range(1,BESS_dim+1):
                mdl.add_constraint(P_DA_CH[(t,j,s)] <= P_max_BESS[s-1] * D_Char[(t,j,s)])                         # 식(20)
                mdl.add_constraint(P_min_BESS[s-1] * D_Char[(t,j,s)] <= P_DA_CH[(t,j,s)])                         # 식(20)
                
                mdl.add_constraint(P_RS_CH[(t,j,s)] <= P_max_BESS[s-1] * D_Char[(t,j,s)] - P_DA_CH[(t,j,s)])     # 식(21)
                mdl.add_constraint(P_min_BESS[s-1] <= P_RS_CH[(t,j,s)])                                          # 식(21)
                
                mdl.add_constraint(P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)] <= P_max_BESS[s-1] * D_Char[(t,j,s)])     # 식(22)
                
                mdl.add_constraint(P_min_BESS[s-1] * D_Char[(t,j,s)] <= P_DA_CH[(t,j,s)] - P_RS_CH[(t,j,s)])     # 식(23)
                
                mdl.add_constraint(P_DA_DCH[(t,j,s)] <= P_max_BESS[s-1] * D_Dchar[(t,j,s)])                      # 식(24)
                mdl.add_constraint(P_min_BESS[s-1] * D_Dchar[(t,j,s)] <= P_DA_DCH[(t,j,s)])                      # 식(24)
                
                mdl.add_constraint(P_RS_DCH[(t,j,s)] <= P_max_BESS[s-1] * D_Dchar[(t,j,s)] - P_DA_DCH[(t,j,s)])  # 식(25)
                mdl.add_constraint(P_min_BESS[s-1] <= P_RS_DCH[(t,j,s)])                                         # 식(25)
                
                mdl.add_constraint(P_UR_CH[(t,j,s)] <= P_RS_CH[(t,j,s)])                                       # 식(26)
                
                mdl.add_constraint(P_DR_CH[(t,j,s)] <= P_RS_CH[(t,j,s)])                                       # 식(27)
                
                mdl.add_constraint(P_UR_DCH[(t,j,s)] <= P_RS_DCH[(t,j,s)])                                     # 식(28)
                
                mdl.add_constraint(P_DR_DCH[(t,j,s)] <= P_RS_DCH[(t,j,s)])                                     # 식(29)
                
                mdl.add_constraint(P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)] <= P_max_BESS[s-1] * D_Dchar[(t,j,s)])  # 식(30)
                
                mdl.add_constraint(P_min_BESS[s-1] * D_Dchar[(t,j,s)] <= P_DA_DCH[(t,j,s)] - P_RS_DCH[(t,j,s)])  # 식(31)
                
                # Energy capacity in the day-ahead and real-time - 식(32)  
                mdl.add_constraint(E_min_BESS[s-1] <= E_BESS[(t,j,s)])  #식(32)
                mdl.add_constraint(E_BESS[(t,j,s)] <= E_max_BESS[s-1])  #식(32)


    ## Capacity of WPR in the dayahead planning - 식(33) ~ 식(36)
    # 식(33)   
    mdl.add_constraints(P_DA_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 
    
    # 식(34)
    mdl.add_constraints(P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] - P_DA_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(34)
    
    # 식(35)
    mdl.add_constraints(P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(35) 

    # 식(36)
    mdl.add_constraints(0 <= P_DA_WPR[(t,j,w)] - P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(36)

    # Deployed power of WPR in the regulation service - 식(37) ~ 식(38)
    mdl.add_constraints(0 <= P_UR_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(37)

    mdl.add_constraints(P_UR_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(37)

    mdl.add_constraints(0 <= P_DR_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(38)

    mdl.add_constraints(P_DR_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(38)   

    ### Constarints of binary decision Variables - 식(39) ~ 식(42)
    ## Commitment status of WPRs, and BESs in the charging and discharging modes in the dayahead planning
    mdl.add_constraints(D_WPR[(t,j,w)] == D_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))       # 식(39)

    mdl.add_constraints(D_Char[(t,j,s)] == D_Char[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(40)

    mdl.add_constraints(D_Dchar[(t,j,s)] == D_Dchar[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(41)

    mdl.add_constraints(0 <= D_Char[(t,j,s)] + D_Dchar[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(42)

    mdl.add_constraints(D_Char[(t,j,s)] + D_Dchar[(t,j,s)] <= 1 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(42)

    ### Constarints of ramp-rate - 식(43) ~ 식(57)
    ## 식(43) and 식(53)
    # t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    # t=1 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ## 식(44) and 식(52)
    # t>=1 and j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1]*del_S <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1]*del_S >= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
 
    # t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))

    ## 식(45) and 식(55)
    # t>=1 and j>=2
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] - (P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] - (P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    # t=1 and j=1  
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ## 식(46) and 식(56)
    # t>=1 and j>=2
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] - (P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] - (P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    # t=1 and j=1 
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ## 식(47)
    # t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1]<= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) - (P_DA_CH[(t,j-1,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) - (P_DA_CH[(t,j-1,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)]) + ( P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)]) + (P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) - (P_DA_CH[(t-1,min_dim,s)] + P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) - (P_DA_CH[(t-1,min_dim,s)] + P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)]) + ( P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)]) + ( P_RS_CH[(t,j,s)]+ P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    # t=1 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))    
    
    ## 식(48)
    # t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) - (P_DA_DCH[(t,j-1,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) - (P_DA_DCH[(t,j-1,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) - (P_DA_DCH[(t-1,min_dim,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) - (P_DA_DCH[(t-1,min_dim,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    # t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))    
    
    ## 식(49) and 식(54)
    # t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))  
    
    # t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)] for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)] for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))

    # t=1 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= P_DA_WPR[(t,j,w)] for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
            
    ## 식(50) and 식(57)
    # t>=1, j>=2
    mdl.add_constraints(Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] - (P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] - (P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,min_dim,w)] for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    # t=1 and j=1
    mdl.add_constraints(Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    ## 식(51)
    # t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) - (P_DA_WPR[(t,j-1,w)] + P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) - (P_DA_WPR[(t,j-1,w)] + P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    #교수님 mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)]) + (P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)]) + ( P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    
    # t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) - (P_DA_WPR[(t-1,min_dim,w)] + P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) - (P_DA_WPR[(t-1,min_dim,w)] + P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    #교수님 mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)] ) + (P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    #교수님 mdl.add_constraints(Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)]) + (P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    # t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
   
    ### Constarints of spillage power - 식(58) ~ 식(59)
    ## 식(58)
    mdl.add_constraints(P_SP_WPR[(t,j,w)] == P_RT_WPR[(t,j,w)] - (P_DA_WPR[(t,j,w)] + P_UR_WPR[(t,j,w)] - P_DR_WPR[(t,j,w)]) 
                        for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))
    
    ## 식(59)
    mdl.add_constraints(P_SP_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] 
                        for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))
                    
    ### Constraints of uncertain parameters-  식(61) ~ 식(63), 논문과 다름
    mdl.add_constraints((1-Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) + mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) <= P_UR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))
    mdl.add_constraints((1+Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) + mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) >= P_UR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))
    mdl.add_constraints((1-Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) + mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) <= P_DR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))
    mdl.add_constraints((1+Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) + mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) >= P_DR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))
    mdl.add_constraints((1-Robust_percent) * Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))
    mdl.add_constraints(P_RT_WPR[(t,j,w)] <= (1+Robust_percent) * Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))
    
    # mdl.add_constraints(0.5 * Expected_P_UR.Cells(t+1,j+1).Value <= P_UR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (61) / 변동구간 +-10%
    # mdl.add_constraints(P_UR[(t,j)] <= 1.5 * Expected_P_UR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (61) / 변동구간 +-10%
    # mdl.add_constraints(0.5 * Expected_P_DR.Cells(t+1,j+1).Value <= P_DR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (62) / 변동구간 +-10%
    # mdl.add_constraints(P_DR[(t,j)] <= 1.5 * Expected_P_DR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (62) / 변동구간 +-10%
    # mdl.add_constraints(0.5 * Expected_P_RT_WPR.Cells(t+1,j+1).Value <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식 (63) / 변동구간 +-10%
    # mdl.add_constraints(P_RT_WPR[(t,j,w)] <= 1.5 * Expected_P_RT_WPR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식 (63) / 변동구간 +-10%

    ### data for excel
    ### Income in day-ahead 전일 수익
    mdl.add_constraints(AV_RO_DA[t] == mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         + Price_RS.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         - mdl.sum(mdl.sum(Marginal_cost_DCH * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH * del_S * P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(Marginal_cost_WPR * del_S * P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) 
                                   for j in range(1,min_dim+1))) for t in range(1,time_dim+1))

    ### B_t - 식(2)
    mdl.add_constraints(B_t[t] == mdl.sum(Price_DA.Cells(t+1,2).Value * del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))
                                          + Price_RS.Cells(t+1,2).Value * del_S * (mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))
                                          + Price_UR.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_UR_DCH[(t,j,s)] - P_UR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))
                                          + Price_DR.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_DR_CH[(t,j,s)] - P_DR_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))
                                          for j in range(1,min_dim+1)) for t in range(1,time_dim+1))   # Income of owner
                     
    ### C_t - 식(3)
    mdl.add_constraints(C_t[t] == mdl.sum(mdl.sum(Marginal_cost_DCH * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH * del_S * P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(Marginal_cost_WPR * del_S * P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1))
                                          + mdl.sum(Marginal_cost_DCH * del_S * (P_UR_DCH[(t,j,s)] + P_DR_DCH[(t,j,s)]) + Marginal_cost_CH * del_S * (P_UR_CH[(t,j,s)] + P_DR_CH[(t,j,s)]) for s in range(1,BESS_dim+1)) + mdl.sum(Marginal_cost_WPR * del_S * P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))
                                          for j in range(1,min_dim+1)) for t in range(1,time_dim+1))   # Cost of owner

    ### Income of BESS#1 in day-ahead
    mdl.add_constraints(BESS1_DA[t] == mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(1,BESS_dim))) for j in range(1,min_dim+1))
                         + Price_RS.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim))) for j in range(1,min_dim+1))
                         - mdl.sum(mdl.sum(Marginal_cost_DCH * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH * del_S * P_DA_CH[(t,j,s)] for s in range(1,BESS_dim)) for j in range(1,min_dim+1))) for t in range(1,time_dim+1))    
    
    ### Income of BESS#2 in day-ahead
    mdl.add_constraints(BESS2_DA[t] == mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(BESS_dim,BESS_dim+1))) for j in range(1,min_dim+1))
                         + Price_RS.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(BESS_dim,BESS_dim+1))) for j in range(1,min_dim+1))
                         - mdl.sum(mdl.sum(Marginal_cost_DCH * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH * del_S * P_DA_CH[(t,j,s)] for s in range(BESS_dim,BESS_dim+1)) for j in range(1,min_dim+1))) for t in range(1,time_dim+1))    
    
    ### Income of WPR in day-ahead
    mdl.add_constraints(WPR_DA[t] == mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * mdl.sum(P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) for j in range(1,min_dim+1))
                         + Price_RS.Cells(t+1,2).Value * mdl.sum(del_S * mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) for j in range(1,min_dim+1))
                         - mdl.sum(mdl.sum(Marginal_cost_WPR * del_S * P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) for j in range(1,min_dim+1))) for t in range(1,time_dim+1))    
    
    ### Income of BESS#1 in real-time
    mdl.add_constraints(BESS1_RT[t] == mdl.sum(Price_UR.Cells(t+1,j+1).Value * del_S * mdl.sum(P_UR_DCH[(t,j,s)] - P_UR_CH[(t,j,s)] for s in range(1,BESS_dim))
                                            + Price_DR.Cells(t+1,j+1).Value * del_S * mdl.sum(P_DR_CH[(t,j,s)] - P_DR_DCH[(t,j,s)] for s in range(1,BESS_dim))
                                            - mdl.sum(Marginal_cost_DCH * del_S * (P_UR_DCH[(t,j,s)] + P_DR_DCH[(t,j,s)]) + Marginal_cost_CH * del_S * (P_UR_CH[(t,j,s)] + P_DR_CH[(t,j,s)]) for s in range(1,BESS_dim)) for j in range(1,min_dim+1)) for t in range(1,time_dim+1))
    
    ### Income of BESS#2 in real-time
    mdl.add_constraints(BESS2_RT[t] == mdl.sum(Price_UR.Cells(t+1,j+1).Value * del_S * mdl.sum(P_UR_DCH[(t,j,s)] - P_UR_CH[(t,j,s)] for s in range(BESS_dim,BESS_dim+1))
                                            + Price_DR.Cells(t+1,j+1).Value * del_S * mdl.sum(P_DR_CH[(t,j,s)] - P_DR_DCH[(t,j,s)] for s in range(BESS_dim,BESS_dim+1))
                                            - mdl.sum(Marginal_cost_DCH * del_S * (P_UR_DCH[(t,j,s)] + P_DR_DCH[(t,j,s)]) + Marginal_cost_CH * del_S * (P_UR_CH[(t,j,s)] + P_DR_CH[(t,j,s)]) for s in range(BESS_dim,BESS_dim+1)) for j in range(1,min_dim+1)) for t in range(1,time_dim+1))
    
    ### Income of WPR in real-time  
    mdl.add_constraints(WPR_RT[t] == mdl.sum(Price_UR.Cells(t+1,j+1).Value * del_S * mdl.sum(P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) 
                                            + Price_DR.Cells(t+1,j+1).Value * del_S * mdl.sum(P_DR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))
                                            - mdl.sum(Marginal_cost_WPR * del_S * P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) for j in range(1,min_dim+1)) for t in range(1,time_dim+1))
    
    return mdl

### Write Result
def result_optimization_model(model, DataFrame):
    mdl = model
    frame = DataFrame
    
    wb_result = excel.Workbooks.Open(os.getcwd()+"\\Result\\robust model_result.xlsx")
    ws1 = wb_result.Worksheets("Optimization Result")
    ws2 = wb_result.Worksheets("Day-Ahead_energy&reserve")
    ws3 = wb_result.Worksheets("Real-time")
    
    ### Sheet 1  
    # Total Revenue
    ws1.Cells(1,2).Value = "Optimization Result"
    ws1.Cells(2,1).Value = "Total Revenue [$]"
    ws1.Cells(2,2).Value = float(mdl.objective_value)
    
    # Income in day-ahead
    ws1.Cells(3,1).Value = "Income in day-ahead [$]"
    ws1.Cells(3,2).Value = frame.loc[frame['var']=="AV-RO-DA"]['index2'].sum()

    # Income in real-time
    ws1.Cells(7,1).Value = "Income in real-time [$]"
    ws1.Cells(7,2).Value = frame.loc[frame['var']=="AV-RO"]['index2'].sum()
    
    # Income of BESS#1 in day-ahead
    ws1.Cells(3,4).Value = "Income of BESS#1 in day-ahead [$]"
    ws1.Cells(3,5).Value = frame.loc[frame['var']=="BESS1-DA"]['index2'].sum()    
    
    # Income of BESS#2 in day-ahead
    ws1.Cells(4,4).Value = "Income of BESS#2 in day-ahead [$]"
    ws1.Cells(4,5).Value = frame.loc[frame['var']=="BESS2-DA"]['index2'].sum()        
    
    # Income of WPR in day-ahead
    ws1.Cells(5,4).Value = "Income of WPR in day-ahead [$]"
    ws1.Cells(5,5).Value = frame.loc[frame['var']=="WPR-DA"]['index2'].sum()       
    
    # Income of BESS#1 in real-time
    ws1.Cells(7,4).Value = "Income of BESS#1 in real-time [$]"
    ws1.Cells(7,5).Value = frame.loc[frame['var']=="BESS1-RT"]['index2'].sum()    
    
    # Income of BESS#2 in real-time
    ws1.Cells(8,4).Value = "Income of BESS#2 in real-time [$]"
    ws1.Cells(8,5).Value = frame.loc[frame['var']=="BESS2-RT"]['index2'].sum()        
    
    # Income of WPR in real-time  
    ws1.Cells(9,4).Value = "Income of WPR in real-time [$]"
    ws1.Cells(9,5).Value = frame.loc[frame['var']=="WPR-RT"]['index2'].sum()    

    # B_t
    ws1.Cells(11,1).Value = "Income of owner [$]"
    ws1.Cells(11,2).Value = frame.loc[frame['var']=="B-t"]['index2'].sum()
    
    # C_t
    ws1.Cells(12,1).Value = "Cost of owner [$]"
    ws1.Cells(12,2).Value = frame.loc[frame['var']=="C-t"]['index2'].sum()  
    
    
    ### Sheet 2
    ws2.Cells(1,1).Value = "var"
    ws2.Cells(1,2).Value = "index1"
    ws2.Cells(1,3).Value = "index2"
    ws2.Cells(1,4).Value = "index3"
    ws2.Cells(1,5).Value = "value"
    P_DA_B_result = frame.loc[frame['var']=="P-DA-B"]['index2'].tolist()
    P_DA_S_result = frame.loc[frame['var']=="P-DA-S"]['index2'].tolist()
    P_DA_CH_result = frame.loc[frame['var']=="P-DA-CH"]['value'].tolist()  
    P_DA_DCH_result = frame.loc[frame['var']=="P-DA-DCH"]['value'].tolist()
    P_DA_WPR_result = frame.loc[frame['var']=="P-DA-WPR"]['value'].tolist()
    P_RS_result = frame.loc[frame['var']=="P-RS"]['index2'].tolist()
    P_RS_CH_result = frame.loc[frame['var']=="P-RS-CH"]['value'].tolist()
    P_RS_DCH_result = frame.loc[frame['var']=="P-RS-DCH"]['value'].tolist()
    P_RS_WPR_result = frame.loc[frame['var']=="P-RS-WPR"]['value'].tolist()
    E_BESS_DA_result = frame.loc[frame['var']=="E-BESS-DA"]['value'].tolist()
    D_Char_result = frame.loc[frame['var']=="D-Char"]['value'].tolist()
    D_Dchar_result = frame.loc[frame['var']=="D-DChar"]['value'].tolist()
    D_WPR_result = frame.loc[frame['var']=="D-WPR"]['value'].tolist()

    for t in range(1,time_dim+1):
        ws2.Cells(t+1,1).Value = "P-DA-B [$]"
        ws2.Cells(t+1,2).Value = t
        ws2.Cells(t+1,3).Value = P_DA_B_result[t-1]
        ws2.Cells(t+25,1).Value = "P-DA-S [$]"
        ws2.Cells(t+25,2).Value = t
        ws2.Cells(t+25,3).Value = P_DA_S_result[t-1] 
        ws2.Cells(t+1489,1).Value = "P-RS"
        ws2.Cells(t+1489,2).Value = t
        ws2.Cells(t+1489,3).Value = P_RS_result[t-1]       
        
        
    for i in range(1,time_dim+1):
        for t in range(1,time_dim+1):
            ws2.Cells(24*(i-1)+t+49,1).Value = "P-DA-CH"
            ws2.Cells(24*(i-1)+t+49,2).Value = i
            ws2.Cells(24*(i-1)+t+49,5).Value = P_DA_CH_result[(i-1)*24 + (t-1)]
            ws2.Cells(24*(i-1)+t+625,1).Value = "P-DA-DCH"
            ws2.Cells(24*(i-1)+t+625,2).Value = i
            ws2.Cells(24*(i-1)+t+625,5).Value = P_DA_DCH_result[(i-1)*24 + (t-1)]           
            ws2.Cells(24*(i-1)+t+1513,1).Value = "P-RS-CH"
            ws2.Cells(24*(i-1)+t+1513,2).Value = i
            ws2.Cells(24*(i-1)+t+1513,5).Value = P_RS_CH_result[(i-1)*24 + (t-1)]                   
            ws2.Cells(24*(i-1)+t+2089,1).Value = "P-RS-DCH"
            ws2.Cells(24*(i-1)+t+2089,2).Value = i
            ws2.Cells(24*(i-1)+t+2089,5).Value = P_RS_DCH_result[(i-1)*24 + (t-1)]                            
            ws2.Cells(24*(i-1)+t+2953,1).Value = "E-BESS-DA"
            ws2.Cells(24*(i-1)+t+2953,2).Value = i
            ws2.Cells(24*(i-1)+t+2953,5).Value = E_BESS_DA_result[(i-1)*24 + (t-1)]  
            ws2.Cells(24*(i-1)+t+3529,1).Value = "D-Char"
            ws2.Cells(24*(i-1)+t+3529,2).Value = i
            ws2.Cells(24*(i-1)+t+3529,5).Value = D_Char_result[(i-1)*24 + (t-1)] 
            ws2.Cells(24*(i-1)+t+4105,1).Value = "D-DChar"
            ws2.Cells(24*(i-1)+t+4105,2).Value = i
            ws2.Cells(24*(i-1)+t+4105,5).Value = D_Dchar_result[(i-1)*24 + (t-1)] 
                                   
            for j in range(1,min_dim+1):
                ws2.Cells(24*(i-1)+j*2+48,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+49,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+624,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+625,3).Value = j                 
                ws2.Cells(24*(i-1)+j*2+1200,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+1201,3).Value = j                  
                ws2.Cells(24*(i-1)+j*2+1512,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+1513,3).Value = j                
                ws2.Cells(24*(i-1)+j*2+2088,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+2089,3).Value = j                                
                ws2.Cells(24*(i-1)+j*2+2952,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+2953,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+3528,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+3529,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+4104,3).Value = j 
                ws2.Cells(24*(i-1)+j*2+4105,3).Value = j 
                                                
                for s in range(1,BESS_dim+1):
                    ws2.Cells(24*(i-1)+j*2+s+47,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+48,4).Value = s 
                    ws2.Cells(24*(i-1)+j*2+s+623,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+624,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+1199,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+1200,4).Value = s  
                    ws2.Cells(24*(i-1)+j*2+s+1511,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+1512,4).Value = s                  
                    ws2.Cells(24*(i-1)+j*2+s+2087,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+2088,4).Value = s                                  
                    ws2.Cells(24*(i-1)+j*2+s+2951,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+2952,4).Value = s                
                    ws2.Cells(24*(i-1)+j*2+s+3527,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+3528,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+4103,4).Value = s
                    ws2.Cells(24*(i-1)+j*2+s+4104,4).Value = s 
   
    for t in range(1,time_dim+1):
        for k in range(1,min_dim+1):
            ws2.Cells(12*(t-1)+k+1201,1).Value = "P-DA-WPR"
            ws2.Cells(12*(t-1)+k+1201,2).Value = t
            ws2.Cells(12*(t-1)+k+1201,5).Value = P_DA_WPR_result[(t-1)*12 + (k-1)]
            ws2.Cells(12*(t-1)+k+2665,1).Value = "P-RS-WPR"
            ws2.Cells(12*(t-1)+k+2665,2).Value = t
            ws2.Cells(12*(t-1)+k+2665,5).Value = P_RS_WPR_result[(t-1)*12 + (k-1)]
            ws2.Cells(12*(t-1)+k+4681,1).Value = "D-WPR"
            ws2.Cells(12*(t-1)+k+4681,2).Value = t
            ws2.Cells(12*(t-1)+k+4681,5).Value = D_WPR_result[(t-1)*12 + (k-1)]  
            
            for j in range(1,min_dim+1):         
                ws2.Cells(12*(t-1)+j+1201,3).Value = j
                ws2.Cells(12*(t-1)+j+2665,3).Value = j
                ws2.Cells(12*(t-1)+j+4681,3).Value = j  
                
                for w in range(1,WPR_dim+1):                   
                    ws2.Cells(12*(t-1)+j+1201,4).Value = w
                    ws2.Cells(12*(t-1)+j+2665,4).Value = w  
                    ws2.Cells(12*(t-1)+j+4681,4).Value = w  

     
    ### Sheet 3
    ws3.Cells(1,1).Value = "var"
    ws3.Cells(1,2).Value = "index1"
    ws3.Cells(1,3).Value = "index2"
    ws3.Cells(1,4).Value = "index3"
    ws3.Cells(1,5).Value = "value"
    P_UR_result = frame.loc[frame['var']=="P-UR"]['index3'].tolist()
    P_UR_DCH_result = frame.loc[frame['var']=="P-UR-DCH"]['value'].tolist()
    P_UR_WPR_result = frame.loc[frame['var']=="P-UR-WPR"]['value'].tolist()  
    P_DR_result = frame.loc[frame['var']=="P-DR"]['index3'].tolist()
    P_DR_CH_result = frame.loc[frame['var']=="P-DR-CH"]['value'].tolist()    
    P_DR_WPR_result = frame.loc[frame['var']=="P-DR-WPR"]['value'].tolist()  
    P_RT_WPR_result = frame.loc[frame['var']=="P-RT-WPR"]['value'].tolist()  
    E_BESS_RT_result = frame.loc[frame['var']=="E-BESS-RT"]['value'].tolist()  
    
    for i in range(1,time_dim+1):
        for t in range(1,time_dim+1):
            ws3.Cells(24*(i-1)+t+289,1).Value = "P-UR-DCH"
            ws3.Cells(24*(i-1)+t+289,2).Value = i
            ws3.Cells(24*(i-1)+t+289,5).Value = P_UR_DCH_result[(i-1)*24 + (t-1)]
            ws3.Cells(24*(i-1)+t+1441,1).Value = "P-DR-CH"
            ws3.Cells(24*(i-1)+t+1441,2).Value = i
            ws3.Cells(24*(i-1)+t+1441,5).Value = P_DR_CH_result[(i-1)*24 + (t-1)]
            ws3.Cells(24*(i-1)+t+2593,1).Value = "E-BESS-RT"
            ws3.Cells(24*(i-1)+t+2593,2).Value = i
            ws3.Cells(24*(i-1)+t+2593,5).Value = E_BESS_RT_result[(i-1)*24 + (t-1)]
                                    
            for j in range(1,min_dim+1):
                ws3.Cells(24*(i-1)+j*2+288,3).Value = j 
                ws3.Cells(24*(i-1)+j*2+289,3).Value = j 
                ws3.Cells(24*(i-1)+j*2+1440,3).Value = j 
                ws3.Cells(24*(i-1)+j*2+1441,3).Value = j 
                ws3.Cells(24*(i-1)+j*2+2592,3).Value = j 
                ws3.Cells(24*(i-1)+j*2+2593,3).Value = j                 
                             
                for s in range(1,BESS_dim+1):
                    ws3.Cells(24*(i-1)+j*2+s+287,4).Value = s
                    ws3.Cells(24*(i-1)+j*2+s+288,4).Value = s     
                    ws3.Cells(24*(i-1)+j*2+s+1439,4).Value = s
                    ws3.Cells(24*(i-1)+j*2+s+1440,4).Value = s      
                    ws3.Cells(24*(i-1)+j*2+s+2591,4).Value = s
                    ws3.Cells(24*(i-1)+j*2+s+2592,4).Value = s 
    
    for t in range(1,time_dim+1):
        for k in range(1,min_dim+1):
            ws3.Cells(12*(t-1)+k+1,1).Value = "P-UR"
            ws3.Cells(12*(t-1)+k+1,2).Value = t
            ws3.Cells(12*(t-1)+k+1,4).Value = P_UR_result[(t-1)*12 + (k-1)]
            ws3.Cells(12*(t-1)+k+865,1).Value = "P-UR-WPR"
            ws3.Cells(12*(t-1)+k+865,2).Value = t
            ws3.Cells(12*(t-1)+k+865,5).Value = P_UR_WPR_result[(t-1)*12 + (k-1)]            
            ws3.Cells(12*(t-1)+k+1153,1).Value = "P-DR"
            ws3.Cells(12*(t-1)+k+1153,2).Value = t
            ws3.Cells(12*(t-1)+k+1153,4).Value = P_DR_result[(t-1)*12 + (k-1)]  
            ws3.Cells(12*(t-1)+k+2017,1).Value = "P-DR-WPR"
            ws3.Cells(12*(t-1)+k+2017,2).Value = t
            ws3.Cells(12*(t-1)+k+2017,5).Value = P_DR_WPR_result[(t-1)*12 + (k-1)] 
            ws3.Cells(12*(t-1)+k+2305,1).Value = "P-RT-WPR"
            ws3.Cells(12*(t-1)+k+2305,2).Value = t
            ws3.Cells(12*(t-1)+k+2305,5).Value = P_RT_WPR_result[(t-1)*12 + (k-1)] 
            
            for j in range(1,min_dim+1):            
                ws3.Cells(12*(t-1)+j+1,3).Value = j
                ws3.Cells(12*(t-1)+j+865,3).Value = j
                ws3.Cells(12*(t-1)+j+1153,3).Value = j
                ws3.Cells(12*(t-1)+j+2017,3).Value = j
                ws3.Cells(12*(t-1)+j+2305,3).Value = j   
                
                for w in range(1,WPR_dim+1):
                    ws3.Cells(12*(t-1)+j+865,4).Value = w
                    ws3.Cells(12*(t-1)+j+2017,4).Value = w
                    ws3.Cells(12*(t-1)+j+2305,4).Value = w     
               
                                     
    print("Optimization Result Calculation Done!")   
    
    wb_result.Save()
    excel.Quit()
    
### Main Program    
if __name__ == '__main__':
    mdl = build_optimization_model() # 최적화 모델 생성
    mdl.print_information() # 모델로부터 나온 정보를 출력
    s = mdl.solve(log_output=True) # 모델 풀기
    
    if s: # 해가 존재하는 경우                
        obj = mdl.objective_value
        mdl.get_solve_details()
        print("* Total cost=%g" % obj)
        print("*Gap tolerance = ", mdl.parameters.mip.tolerances.mipgap.get())
        
        data = [v.name.split('_') + [s.get_value(v)] for v in mdl.iter_variables()] # 변수 데이터 저장
        frame = pd.DataFrame(data, columns=['var', 'index1', 'index2', 'index3', 'value']) # 변수 중 시간 성분만 있는 경우 'index2'에 값이 저장됨
        frame.to_excel(os.getcwd()+"\\Result\\variable_result.xlsx")
        
        result_optimization_model(mdl, frame)  # 결과 출력부        
        
        # Save the CPLEX solution as "solution.json" program output
        with get_environment().get_output_stream("solution.json") as fp: #json 형태로 solution 저장
            mdl.solution.export(fp, "json")
        
    else: # 해가 존재하지 않는 경우
        print("* model has no solution")
        excel.Quit() 
    