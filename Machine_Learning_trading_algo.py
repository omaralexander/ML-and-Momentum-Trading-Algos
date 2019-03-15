from quantopian.pipeline import Pipeline, CustomFilter,CustomFactor
from quantopian.algorithm import attach_pipeline, pipeline_output,order_optimal_portfolio 
from quantopian.pipeline.factors import Latest
from quantopian.pipeline.data.builtin import USEquityPricing
from sklearn.ensemble import RandomForestClassifier
from quantopian.pipeline.factors import SimpleMovingAverage,ExponentialWeightedMovingAverage
from quantopian.pipeline.filters import Q1500US
from quantopian.pipeline.data import morningstar
import quantopian.optimize as opt

import numpy as np
import pandas as pd
from scipy import stats

momentum_length_1 = 2     # 20
momentum_length_2 = 5     # 60
momentum_length_3 = 8    # 125

def _slope(ts):
    x = np.arange(len(ts))
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, ts)
    annualized_slope = (1 + slope)**250
    return annualized_slope * (r_value ** 2)   

class Momentum_factors(CustomFactor):
    inputs = [USEquityPricing.close]
    window_length = 10 # 260
    
    def compute(self,today,assets,out,close):
        
        value_table           = pd.DataFrame(index=assets)
        value_table['first']  = close[-1] / close[momentum_length_1] - 1
        value_table['second'] = close[-1]/ close[momentum_length_2] - 1
        value_table['third']  = close[-1] / close[momentum_length_3] - 1
        value_table['fourth'] = close[-1] / close[0] -1
        
        out[:] = value_table.rank(ascending = False).mean(axis=1)
        
class stddev_factor(CustomFactor):
    inputs  = [USEquityPricing.close]
    window_length = 5
    
    def compute(self,today,assets,out,close):
        out[:] = np.nanstd(np.diff(close, axis=0), axis = 0)
                 
class current_ratio(CustomFactor):
    inputs = [morningstar.operation_ratios.current_ratio]
    
    def compute(self,today,assets,out,ratio):
        out[:] = np.mean(ratio, axis = 0)
        
 
def initialize(context):

    context.securities_in_results  = []   
    context.model                  = RandomForestClassifier()
    context.model.n_estimators     = 400
    context.model.min_samples_leaf = 200
    context.lookback               = 3 # Look back 3 days
    context.history_range          = 15
    context.max_price              = 100
    context.min_price              = 3

    attach_pipeline(Custom_pipeline(context), 'Custom_pipeline')
    schedule_function(trade, date_rules.every_day(), time_rules.market_open())
    schedule_function(sell,  date_rules.every_day(), time_rules.market_open())
    schedule_function(buy,   date_rules.every_day(), time_rules.market_open(minutes = 5))
    #schedule_function(park,date_rules.every_day(),time_rules.market_open(minutes = 30))
    
def Custom_pipeline(context):

    pipe       = Pipeline()
    curr_price = USEquityPricing.close.latest
    
    my_universe = Q1500US()&(context.min_price < curr_price)&(curr_price < context.max_price)

    current_filter = current_ratio(window_length = 30, mask = my_universe)
    
    ema_10 = ExponentialWeightedMovingAverage(  
                            inputs        = [USEquityPricing.close],  
                            window_length = 30,  
                            decay_rate    = (1 - (2.0 / (1 + 15.0))),
                            mask          = current_filter > 1.5) 
    
    close           = USEquityPricing.close
    exp_filter      = close > ema_10
    stocks          = exp_filter
    total_filter    = (stocks)&(context.min_price < curr_price)&(curr_price < context.max_price)
    
    m_factor_1      = Momentum_factors(mask = total_filter)
    m_factor_3      = stddev_factor(mask = total_filter)
    #m_factor_3_rank = m_factor_3.rank(mask = total_filter, ascending=True)
    m_factor_3_rank = m_factor_1.rank(mask = total_filter, ascending=True)
    combo_raw       = ((m_factor_3_rank * 0.5))/1
    
    #pipe.add(m_factor_1,'momentum')
    pipe.add(m_factor_3,'factor_3')
    pipe.add(combo_raw,'combo_raw')
    
    pipe.add(combo_raw.rank(mask=total_filter),'combo_rank')
    pipe.set_screen(total_filter)
    
    return pipe

def before_trading_start(context, data):

    context.output = pipeline_output('Custom_pipeline').dropna()
    stocks         = context.output.sort(['combo_rank'],ascending=True)
    
    context.securities_in_results = []
    
    for s in stocks.iloc[:20].index: context.securities_in_results.append(s) 
    if len(context.securities_in_results) > 0.0: log.info(context.output)
        
def trade (context, data):
    context.longs=[]

    if len(context.securities_in_results) > 0.0:                    
        for sec in context.securities_in_results:
           
            recent_prices  = data.history(sec, 'price', context.history_range, '1d').values 
            recent_volumes = data.history(sec, 'volume', context.history_range, '1d').values
            price_changes  = np.diff(recent_prices).tolist() 
            volume_changes = np.diff(recent_volumes).tolist()

            X,Y = [],[]
            
            for i in range(context.history_range-context.lookback-1): 
                X.append(price_changes[i:i+context.lookback] + volume_changes[i:i+context.lookback])
                Y.append(price_changes[i+context.lookback])
                
            context.model.fit(X, Y) 
            
            recent_prices  = data.history(sec, 'price', context.lookback+1, '1d').values 
            recent_volumes = data.history(sec, 'volume',context.lookback+1, '1d').values
            price_changes  = np.diff(recent_prices).tolist() 
            volume_changes = np.diff(recent_volumes).tolist()
            prediction     = context.model.predict(price_changes + volume_changes)


            if prediction > 2.0:#2.0 was the best 
                if sec not in context.portfolio.positions: context.longs.append(sec)
                
def buy (context, data):
    if len(context.longs) != 0: order_target_percent(sid(41382),0)
    for sec in context.longs: order_target_percent(sec, 1.0/len(context.longs))
        
        
def sell (context, data):
    for sec in context.portfolio.positions:
        if sec not in context.longs:
            order_target_percent(sec, 0.0)

    if len(context.portfolio.positions) == 0:
        safe_stock   = sid(41382)
        price        = data.current(safe_stock,'price')
        order_target(safe_stock, 1, style=LimitOrder(price))
        
            
def park(context,data):
    
    safe_stock   = sid(41382)
    price        = data.current(safe_stock,'price')
    amount       = context.portfolio.cash // price
    count        = len(context.portfolio.positions)
    in_portfolio = False
        
    try:             in_portfolio = safe_stock in context.portfolio.positions
    except KeyError: in_portfolio = False

    if count != 0 and in_portfolio: order_target_percent(safe_stock,0)
    elif count == 0 and in_portfolio == False: order_target(safe_stock, 1, style=LimitOrder(price))
        
def handle_data (context, data):
    record(leverage  = context.account.leverage,
           positions = len(context.portfolio.positions),
           resutls   = len(context.securities_in_results))





