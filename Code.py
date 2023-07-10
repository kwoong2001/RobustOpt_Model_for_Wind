# 23.07.07~
# 변경 프로젝트
#   - DA->DA
#   - RS->DA_RS
#   - UR->RT_UR_imbalance (RT-LMP) 
#   - DR->RT_DR_imbalance (RT-LMP) 

from __future__ import print_function
from cmath import inf
from docplex.mp.model import Model
from docplex.util.environment import get_environment
import win32com.client as win32
import pandas as pd
import os

excel = win32.Dispatch("Excel.Application")
wb1 = excel.Workbooks.Open(os.getcwd()+"\\Data\\robust model_data_modified.xlsx")
Price_DA = wb1.Sheets("Price_DA")              # Day-ahead price
Price_RS = wb1.Sheets("Price_RS")              # Reserve prices
Price_UR = wb1.Sheets("Price_UR")              # Up regulation prices
Price_DR = wb1.Sheets("Price_DR")              # Down regulation prices
Expected_P_RT_WPR = wb1.Sheets("Expected_P_RT_WPR")  # Expected wind power realization

### 파라미터 설정
time_dim = 24     # 시간 개수 (t)
min_dim = 12      # ex) 5분 x 12 = 1시간 (j)
del_S = 1/min_dim # Duration of intra-hourly interval ex) 5min = 1/12(h)
BESS_dim = 2      # BESS 개수 (s)
WPR_dim = 1       # 풍력발전기 개수 (w)
Marginal_cost_CH = [1,0.8]    # Marginal cost of BES in charging modes
Marginal_cost_DCH = [1,0.8]   # Marginal cost of BES in discharging modes
Marginal_cost_WPR = [3]   # Marginal cost of WPR
Ramp_rate_WPR = 3       # Ramp-rate of WPR (MW)
E_min_BESS = [0,0]    # Minimum energy of BESS (MWh)
E_max_BESS = [30,18]   # Maximum energy of BESS (MWh)
P_max_BESS = [5,3]    # Maximum power of BESS (MW)
P_min_BESS = [0,0]    # Minimum power of BESS (MW)
Ramp_rate_BESS = [5,3]  # Ramp-rate of BESS

Robust_percent = 0.5   # Robust percent (0~1)
contri_reg_percent = 0.5 # Up regulation 혹은 down regulation에 기여하는 비율 (0~1)

### 최적화 파트
def build_optimization_model(name='Robust_Optimization_Model'):
    mdl = Model(name=name)   # Model - Cplex에 입력할 Model 이름 입력 및 Model 생성
    mdl.parameters.mip.tolerances.mipgap = 0.0001;   # 최적화 계산 오차 설정

    time = [t for t in range(1,time_dim+1)]    # (t)의 one dimension
    time_min = [(t,j) for t in range(1,time_dim + 1) for j in range(1,min_dim+1)]   # (t,j)의 two dimension
    time_n_BESS = [(t,j,s) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1)]   # (t,j,s)의 three dimension
    time_n_WPR = [(t,j,w) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1)]     # (t,j,w)의 three dimension

    ### Continous Variable 지정 (연속 변수, 실수 변수)
    #Day-ahead
    P_DA_S = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-S")   # Selling bids in the day-ahead market
    P_DA_B = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-B")   # Buying bids in the day-ahead market
    P_RS = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-RS")       # Reserve bid 
    
    P_UR = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-UR")   # Deployed power in the up-regulation services
    P_DR = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-DR")   # Deployed power in the down-regulation services

    P_DA_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-CH")     # Day-ahead scheduling of BES in charging modes
    P_DA_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-DCH")   # Day-ahead scheduling of BES in discharging modes
    P_DA_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DA-WPR")    # Day-ahead scheduling of WPR

    P_UR_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-UR-CH")     # Deployed up regulation power of BES in charging mode
    P_UR_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-UR-DCH")   # Deployed up regulation power of BES in discharging mode
    P_UR_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-UR-WPR")    # Deployed up regulation power of WPR   

    P_DR_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DR-CH")      # Deployed down regulation power of BES in charging mode
    P_DR_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DR-DCH")    # Deployed down regulation power of BES in discharging mode
    P_DR_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DR-WPR")     # Deployed down regulation power of WPR  

    P_RS_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-CH")      # Reserve scheduling of BES in charging modes
    P_RS_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-DCH")    # Reserve scheduling of BES in discharging modes
    P_RS_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RS-WPR")     # Reserve scheduling of WPR

    #Real-time
    P_SP_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-SP-WPR")            # Spilled power of WPR (difference between the realization of wind power and the scheduled power of WPR)
    E_BESS = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="E-BESS")        # Energy level of BES in Real-time
    P_RT_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RT-WPR")                # Realization of wind power in real-time

    AV_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="AV-WPR")                    # Auxiliary variables for linearization
    
    
    ### Functions
    AV_RO = mdl.continuous_var_dict(time, lb=0, ub=inf, name="AV-RO")                        # Auxiliary variable of RO
    B_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="B-t")          # Income function of owner
    C_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="C-t")          # Cost function of owner

    ### Binary Variable 지정 (이진 변수)
    D_Char = mdl.binary_var_dict(time_n_BESS, name="D-Char-DA")      # Charging binary variables of BES (알파)
    D_Dchar = mdl.binary_var_dict(time_n_BESS, name="D-DChar-DA")    # Discharging binary variables of BES (베타)
    D_WPR = mdl.binary_var_dict(time_n_WPR, name="D-WPR")         # Commitment status binary variable of WPR
    
    ### Objective function - 식(1) / 식(65)
    
    mdl.maximize(mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         + Price_RS.Cells(t+1,2).Value * mdl.sum(del_S * ((mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))) for j in range(1,min_dim+1))
                         - mdl.sum((mdl.sum(Marginal_cost_DCH[s-1] * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH[s-1] * del_S * P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(Marginal_cost_WPR[w-1] * del_S * P_DA_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         + AV_RO[t] for t in range(1,time_dim+1)))

    # Robust Optizimation을 위한 변수 (BESS + WPR) - 식(65)
    #original
    mdl.add_constraints(AV_RO[t] <= mdl.sum(Price_UR.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_UR_DCH[(t,j,s)] - P_UR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) 
                                            + Price_DR.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_DR_CH[(t,j,s)]- P_DR_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DR_WPR[(t,j,w)] for w in range(1,WPR_dim+1)))
                                            - mdl.sum(Marginal_cost_DCH[s-1] * del_S * (P_UR_DCH[(t,j,s)] + P_DR_DCH[(t,j,s)])  + Marginal_cost_CH[s-1] * del_S * (P_UR_CH[(t,j,s)] + P_DR_CH[(t,j,s)])  for s in range(1,BESS_dim+1)) - mdl.sum(Marginal_cost_WPR[w-1] * del_S * P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))
                                            - mdl.sum(Price_DR.Cells(t+1,j+1).Value * del_S * P_SP_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) 
                                            for j in range(1,min_dim+1)) for t in range(1,time_dim+1))
    
    ### Equality constraints - 식(4) ~ 식(6) + 식(12) ~ 식(14)
    #Day-ahead bid 식(4)~식(6)
    mdl.add_constraints(P_DA_DCH[(t,j,s)] == P_DA_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(4)

    mdl.add_constraints(P_DA_CH[(t,j,s)] == P_DA_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(5)

    mdl.add_constraints(P_DA_WPR[(t,j,w)] == P_DA_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(6

    #Reserve bid 식(12)~식(14)
    mdl.add_constraints(P_RS_CH[(t,j,s)] == P_RS_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(12)   

    mdl.add_constraints(P_RS_DCH[(t,j,s)] == P_RS_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(13)                                           

    mdl.add_constraints(P_RS_WPR[(t,j,w)] == P_RS_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(14)

    ### Constraints of day-ahead energy / reserve bids / real-time deployed power in the up and down regulation services - 식(7) ~ 식(11), 
    
    #식(7)-(9)
    mdl.add_constraints(P_DA_S[t] == mdl.sum(mdl.sum(P_DA_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum( P_DA_WPR[(t,j,w)]  for w in range(1,WPR_dim+1) )) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # 식(7)

    mdl.add_constraints(P_DA_B[t] == mdl.sum(P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # 식(8)
    
    mdl.add_constraints(P_RS[t] == mdl.sum(mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_WPR[(t,j,w)] for w in range(1,WPR_dim+1)) ) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # 식(9)
    
    #식(10)-(11)
    #original
    mdl.add_constraints(P_UR[(t,j)] == mdl.sum(mdl.sum(P_UR_DCH[(t,j,s)] + P_UR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_UR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # 식(10)
    mdl.add_constraints(P_DR[(t,j)] == mdl.sum(mdl.sum(P_DR_DCH[(t,j,s)] + P_DR_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DR_WPR[(t,j,w)] for w in range(1,WPR_dim+1))) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # 식(11)
    
    ### 식(15) ~ 식(16)
    mdl.add_constraints(P_UR[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(15)
    mdl.add_constraints(P_DR[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(16)
    #mdl.add_constraints(P_DR[(t,j)] <= 0 for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(16)
    
    ### Constarints of stored energy of BES - 식(17) ~ 식(19)
    ## Day-ahead
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
    # Power capacity of BES in day-ahead planning - 식(20) ~ 식(25) + 식(30) ~ 식(31)
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
                
                mdl.add_constraint(E_min_BESS[s-1] <= E_BESS[(t,j,s)])  #식(32)
                mdl.add_constraint(E_BESS[(t,j,s)] <= E_max_BESS[s-1])  #식(32)
                
    # Energy capacity in the real-time - 식(32)
    
    # Capacity of WPR in the dayahead planning - 식(33) ~ 식(36)
    #식(33)
   
    mdl.add_constraints(P_DA_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)]  for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 
    
    #식(34)
    mdl.add_constraints(P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] - P_DA_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(34)

    #식(35)
    mdl.add_constraints(P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(35)
    
    #식(36)
    mdl.add_constraints(0 <= P_DA_WPR[(t,j,w)] - P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(36)

    # Deployed power of WPR in the regulation service - 식(37) ~ 식(38)
    mdl.add_constraints(0 <= P_UR_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(37)

    mdl.add_constraints(P_UR_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(37)

    mdl.add_constraints(0 <= P_DR_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(38)

    mdl.add_constraints(P_DR_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(38)   

    ### Constarints of binary decision Variables - 식(39) ~ 식(42)
    # Commitment status of WPRs, and BESs in the charging and discharging modes in the dayahead planning
    mdl.add_constraints(D_WPR[(t,j,w)] == D_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))       # 식(39)

    mdl.add_constraints(D_Char[(t,j,s)] == D_Char[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(40)

    mdl.add_constraints(D_Dchar[(t,j,s)] == D_Dchar[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(41)

    mdl.add_constraints(0 <= D_Char[(t,j,s)] + D_Dchar[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(42)

    mdl.add_constraints(D_Char[(t,j,s)] + D_Dchar[(t,j,s)] <= 1 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(42)
    
    ### Constarints of ramp-rate - 식(43) ~ 식(57)
    ##식(43) and 식(53)
    #t>=1, j>=2 - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1  - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##식(44) and 식(52)
    #t>=1 and j>=2 - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1 ,  - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##식(45) and 식(55), 논문이 맞음, 예비력은 사용되지 않았기 때문에 차이로 Ramp rate를 계산할 수 없음
    #t>=1 and j>=2,
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 , 
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1 ,
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##식(46) and 식(56)
    #t>=1 and j>=2,
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 , 
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1 
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##식(47)
    #t>=1, j>=2,   - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)]) + ( P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)]) + (P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 , 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)]) + ( P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)]) + ( P_RS_CH[(t,j,s)]+ P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1 , 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##식(48)
    #t>=1, j>=2,   - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints( Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##식(49) and 식(54)
    #t>=1, j>=2,   - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_WPR  <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    mdl.add_constraints( Ramp_rate_WPR >= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    
    #t>=2 and j=1 , 
    mdl.add_constraints(-1 * Ramp_rate_WPR  <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)] for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints( Ramp_rate_WPR >= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)] for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    #t=1 and j=1 , 
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints( Ramp_rate_WPR >= P_DA_WPR[(t,j,w)] for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    ##식(50) and 식(57), 논문이 맞음, 예비력은 사용되지 않았기 때문에 차이로 Ramp rate를 계산할 수 없음
    #t>=1, j>=2, 
    mdl.add_constraints( Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    
    #t>=2 and j=1 , 
    mdl.add_constraints( Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,min_dim,w)] for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    #t=1 and j=1 , 
    mdl.add_constraints( Ramp_rate_WPR >= P_RS_WPR[(t,j,w)] for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    ##식(51)
    #t>=1, j>=2,   - 필요없음, 같은 t에서 j에 관계없이 일정한 값을 가짐
    mdl.add_constraints(-1 * Ramp_rate_WPR<= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)]) + (P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    mdl.add_constraints( Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)]) + ( P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for w in range(1,WPR_dim+1))
    
    #t>=2 and j=1 , 
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)] ) + (P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints( Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,min_dim,w)]) + (P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,min_dim,w)]) for t in range(2,time_dim+1) for j in range(1,2) for w in range(1,WPR_dim+1))
    
    #t=1 and j=1 ,
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
    mdl.add_constraints( Ramp_rate_WPR >= (P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)]) for t in range(1,2) for j in range(1,2) for w in range(1,WPR_dim+1))
       
    ### Constarints of spillage power - 식(58) ~ 식(59)
    ##식(58)
    mdl.add_constraints(P_SP_WPR[(t,j,w)] == P_RT_WPR[(t,j,w)] - (P_DA_WPR[(t,j,w)] + P_UR_WPR[(t,j,w)] - P_DR_WPR[(t,j,w)]) 
                        for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))
    
    ##식(59)
    mdl.add_constraints(P_SP_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] 
                        for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))

                        
    ### Constraints of uncertain parameters-  식(61) ~ 식(63), 논문과 다름
    mdl.add_constraints((1-Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1))+ mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) <= P_UR[(t,j)] for s in range(1, BESS_dim+1) for t in range(1,time_dim+1) for j in range(1,min_dim+1))

    mdl.add_constraints((1+Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1))+ mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) >= P_UR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (61) / 변동구간 +-10%

    mdl.add_constraints((1-Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1))+ mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) <= P_DR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (61) / 변동구간 +-10%

    mdl.add_constraints((1+Robust_percent) * contri_reg_percent * (mdl.sum(Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for w in range(1,WPR_dim+1))+ mdl.sum(P_max_BESS[s-1] for s in range(1,BESS_dim+1))) >= P_DR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (61) / 변동구간 +-10%

    mdl.add_constraints((1-Robust_percent) * Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식 (63) / 변동구간 +-10%

    mdl.add_constraints(P_RT_WPR[(t,j,w)] <= (1+Robust_percent) * Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식 (63) / 변동구간 +-10%

    return mdl

### Write Result
def result_optimization_model(model, DataFrame):
    mdl = model
    frame = DataFrame
    
    wb_result = excel.Workbooks.Open(os.getcwd()+"\\Result\\robust model_result.xlsx")
    ws1 = wb_result.Worksheets("Optimization Result")
    ws2 = wb_result.Worksheets("Day-Ahead")
    
    ### Sheet 1  
    # Total Revenue
    ws1.Cells(1,2).Value = "Optimization Result"
    ws1.Cells(2,1).Value = "Total Revenue [$]"
    ws1.Cells(2,2).Value = float(mdl.objective_value)
       
    # AV_RO
    ws1.Cells(3,1).Value = "Income in real-time [$]"
    ws1.Cells(3,2).Value = frame.loc[frame['var']=="AV-RO"]['index2'].sum()
    
    ### Case Result
    ws1.Cells(7,1).Value = "Contribution"
    
    ## Day-ahead Result
    ws1.Cells(5,2).Value = "Day-head"
    
    #Result for BESS
    for s in range(1, BESS_dim+1):
        if s == 1:
            BESS_DA_DCH_cost_result = []
            BESS_DA_CH_cost_result = []
            BESS_DA_DCH_price_result = []
            BESS_DA_CH_price_result = []
            BESS_DA_RS_price_result=[]
            
        BESS_DA_DCH_cost_result.append(0)
        BESS_DA_CH_cost_result.append(0)
        BESS_DA_DCH_price_result.append(0)
        BESS_DA_CH_price_result.append(0)
        BESS_DA_RS_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                BESS_DA_DCH_cost_result[s-1] += del_S * Marginal_cost_DCH[s-1] * frame.loc[frame['var']=="P-DA-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                BESS_DA_CH_cost_result[s-1] += del_S * Marginal_cost_CH[s-1] * frame.loc[frame['var']=="P-DA-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    
                BESS_DA_DCH_price_result[s-1] += del_S * Price_DA.Cells(t+1,2).Value * frame.loc[frame['var']=="P-DA-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                BESS_DA_CH_price_result[s-1] += del_S * Price_DA.Cells(t+1,2).Value * frame.loc[frame['var']=="P-DA-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                
                BESS_DA_RS_price_result[s-1] += del_S * Price_RS.Cells(t+1,2).Value * (frame.loc[frame['var']=="P-RS-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                                                       +frame.loc[frame['var']=="P-RS-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                    
        ws1.Cells(6,2+s-1).Value = "BESS#"+str(s)
        ws1.Cells(7,2+s-1).Value = BESS_DA_RS_price_result[s-1]+(BESS_DA_DCH_price_result[s-1]-BESS_DA_CH_price_result[s-1])-(BESS_DA_DCH_cost_result[s-1] + BESS_DA_CH_cost_result[s-1])
        
    #Result for Wind
    for w in range(1, WPR_dim+1):
        if w == 1:
            Wind_DA_cost_result = []
            Wind_DA_price_result = []
            Wind_DA_RS_price_result=[]
        Wind_DA_cost_result.append(0)
        Wind_DA_price_result.append(0)
        Wind_DA_RS_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                Wind_DA_cost_result[w-1] += del_S * Marginal_cost_WPR[w-1] * frame.loc[frame['var']=="P-DA-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum()
                Wind_DA_price_result[w-1] += del_S * Price_DA.Cells(t+1,2).Value * frame.loc[frame['var']=="P-DA-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum()
                Wind_DA_RS_price_result[w-1] += del_S * Price_RS.Cells(t+1,2).Value * frame.loc[frame['var']=="P-RS-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum()
                    
        ws1.Cells(6,2+BESS_dim+w-1).Value = "Wind#"+str(w)
        ws1.Cells(7,2+BESS_dim+w-1).Value = Wind_DA_RS_price_result[w-1]+Wind_DA_price_result[w-1]-Wind_DA_cost_result[w-1]
    
    ## Real-time Result
    ws1.Cells(5,2+BESS_dim+WPR_dim).Value = "Real-time"
    #Result for BESS
    for s in range(1, BESS_dim+1):
        if s == 1:
            BESS_RT_DCH_cost_result = []
            BESS_RT_CH_cost_result = []
            BESS_RT_DCH_price_result = []
            BESS_RT_CH_price_result = []
        BESS_RT_DCH_cost_result.append(0)
        BESS_RT_CH_cost_result.append(0)
        BESS_RT_DCH_price_result.append(0)
        BESS_RT_CH_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                
                BESS_RT_DCH_cost_result[s-1] += del_S * Marginal_cost_DCH[s-1] * (frame.loc[frame['var']=="P-UR-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                                          + frame.loc[frame['var']=="P-DR-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()) 
                BESS_RT_CH_cost_result[s-1] += del_S * Marginal_cost_CH[s-1] * (frame.loc[frame['var']=="P-UR-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                                        +frame.loc[frame['var']=="P-DR-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                    
                BESS_RT_DCH_price_result[s-1] += (Price_UR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-UR-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                    - Price_DR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-DR-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                BESS_RT_CH_price_result[s-1] += (Price_UR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-UR-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                    - Price_DR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-DR-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                    
        ws1.Cells(6,2+BESS_dim+WPR_dim+s-1).Value = "BESS#"+str(s)
        ws1.Cells(7,2+BESS_dim+WPR_dim+s-1).Value = (BESS_RT_DCH_price_result[s-1]-BESS_RT_CH_price_result[s-1])-(BESS_RT_DCH_cost_result[s-1] + BESS_RT_CH_cost_result[s-1])
        
    #Result for Wind
    for w in range(1, WPR_dim+1):
        if w == 1:
            Wind_RT_cost_result = []
            Wind_RT_price_result = []
        Wind_RT_cost_result.append(0)
        Wind_RT_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                Wind_RT_cost_result[w-1] += del_S * Marginal_cost_WPR[w-1] * frame.loc[frame['var']=="P-UR-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum()
                Wind_RT_price_result[w-1] += (Price_UR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-UR-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum()
                                                +Price_DR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-DR-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum()
                                                -Price_DR.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-SP-WPR"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+w].sum())
                    
        ws1.Cells(6,2+BESS_dim+WPR_dim+BESS_dim+w-1).Value = "Wind#"+str(w)
        ws1.Cells(7,2+BESS_dim+WPR_dim+BESS_dim+w-1).Value = Wind_RT_price_result[w-1]-Wind_RT_cost_result[w-1]    
    
    ### Sheet 2
    
          
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
    