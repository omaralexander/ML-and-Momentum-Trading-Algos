"""
This is a momentum based strategy that sorts by the close and the stdDev amount
"""
from quantopian.algorithm import attach_pipeline, pipeline_output 
from quantopian.pipeline.experimental import risk_loading_pipeline
from quantopian.pipeline import Pipeline
from quantopian.algorithm import order_optimal_portfolio
from quantopian.pipeline import CustomFactor  
from quantopian.pipeline.data.builtin import USEquityPricing 
from quantopian.pipeline.filters import QTradableStocksUS
from quantopian.pipeline.factors import RSI
from quantopian.pipeline.factors import AverageDollarVolume
from quantopian.pipeline.factors import RollingLinearRegressionOfReturns
from quantopian.pipeline.data import morningstar


import numpy as np
import talib as ta
import pandas as pd
import scipy as sp
import quantopian.optimize as opt
from collections import defaultdict
import datetime
import pytz


class momentum_factor_1(CustomFactor):    
   inputs = [USEquityPricing.close]   
   window_length = 20  
     
   def compute(self, today, assets, out, close):      
     out[:] = close[-1]/close[0]      

class momentum_factor_2(CustomFactor):
    inputs = [USEquityPricing.close]
    window_length = 260
    
    def compute(self,today,assets,out,close):
        out[:] = close[-1]/close[0]

class stddev_factor(CustomFactor):
    inputs = [USEquityPricing.close]
    window_length = 260
    
    def compute(self,today,assets,out,close):
        out[:] = np.nanstd(np.diff(close, axis=0), axis = 0)

            
class market_cap(CustomFactor):    
   inputs = [USEquityPricing.close, morningstar.valuation.shares_outstanding]   
   window_length = 1  
     
   def compute(self, today, assets, out, close, shares):      
     out[:] = close[-1] * shares[-1]     


stock_purchase_date = []

def initialize(context):
    context.max_pos_size = 0.01
    context.max_turnover = 0.95
    context.holdings = 10
    context.profit_taking = 20
    context.max_loss = -20
    context.max_leverage = 1
    set_slippage(slippage.FixedSlippage(spread=0.00))
    
    schedule_function(monthly_rebalance,
                     date_rules.month_end(),
                     time_rules.market_open(hours = 5)) #5 is 8%

    pipe = Pipeline()
    
    attach_pipeline(pipe,'ranked_stocks')
    
    attach_pipeline(risk_loading_pipeline(),'risk_pipe')
            
    market_cap_filter = market_cap()
    stocks = market_cap_filter.top(3000)
    total_filter = (stocks)
    pipe.set_screen(total_filter)
    
    m_factor1 = momentum_factor_1()
    pipe.add(m_factor1,'factor_1')
    m_factor2 = momentum_factor_2()
    pipe.add(m_factor2,'factor_2')
    m_factor_3 = stddev_factor()
    pipe.add(m_factor_3,'factor_3')


    m_factor1_rank = m_factor1.rank(mask = total_filter, ascending = False)
    m_factor2_rank = m_factor2.rank(mask = total_filter, ascending = False)
    m_factor_3_rank = m_factor_3.rank(mask = total_filter, ascending=True)

    
    pipe.add(m_factor1_rank,'f1_rank')
    pipe.add(m_factor2_rank,'f2_rank')
    pipe.add(m_factor_3_rank,'f3_rank')


    #should put weighted adding instead
    combo_raw = (m_factor1_rank + m_factor2_rank +m_factor_3_rank)/3
    pipe.add(combo_raw,'combo_raw')
    pipe.add(combo_raw.rank(mask=total_filter),'combo_rank')
    
        
def monthly_rebalance(context,data):
    
    context.output = pipeline_output('ranked_stocks').dropna()
    context.risk_factor_betas = pipeline_output('risk_pipe')
    context.stocks_ranked = context.output.sort(['combo_rank'],ascending=True).iloc[:context.holdings]
    context.output['date'] = get_datetime().date()
    context.stocks = context.stocks_ranked
    context.output['weight'] = 0.05
    price = data.history(sid(8554),'price',16,'1d')
    spy_rsi = ta.RSI(price,timeperiod = 15)[-1]
    record(current_holdings = len(context.portfolio.positions), leverage = context.account.leverage)


    for stock in context.stocks.index:

        if stock not in context.portfolio.positions and stock not in get_open_orders():
            target_long_to_order = context.portfolio.cash / context.max_leverage
            target_price = data.current(stock,'price')
            target_shares = target_long_to_order // target_price

            #order(stock,target_shares,style=LimitOrder(target_price))
            context.stocks.loc[stock,'weight'] = 0.15
            
            
            #stock_purchase_date.append([stock,get_datetime().date()])

        else:
             price_paid = context.portfolio.positions[stock].cost_basis
             current_shares = context.portfolio.positions[stock].amount
             current_pnl= price_paid * current_shares
             
             stock_pnl = context.portfolio.positions[stock].last_sale_price * current_shares
             delta = safe_div(stock_pnl,current_pnl)
             
             if delta > context.profit_taking:
                
                #order_target_percent(stock,0)
                context.stocks.loc[stock,'weight'] = 0
    objective = opt.TargetWeights(context.stocks.weight.dropna())
    max_thing = opt.MaxGrossExposure(context.max_leverage)
    order_optimal_portfolio(objective = objective,constraints=[max_thing])  
    

                                                   
def find_stock_location(element,list_element):
    try:
        location = [location for location in list_element if element in location][0]
        return location
    except IndexError:
        return 'null'
    

def safe_div(x,y):
    try: return ((x-y)/x) * 100
    except ZeroDivisionError: return 0