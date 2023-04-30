from BotConfig import *
from BotKIKr import BotKIKr
import pandas as pd
import numpy as np
import pickle
import requests
import datetime
import threading
import math
import os

class Bot_15m():

    
    def __init__(self):

        self.mock = True
        self.key = KI_APPKEY_IMITATION if self.mock else KI_APPKEY_PRACTICE
        self.secret = KI_APPSECRET_IMITATION if self.mock else KI_APPSECRET_PRACTICE
        self.account = KI_ACCOUNT_IMITATION if self.mock else KI_ACCOUNT_PRACTICE

        self.bkk = BotKIKr(self.key, self.secret, self.account, self.mock)
        self.bdf = None
        self.b_l = None
        self.q_l = None
        self.r_l = None

        self.tot_evl_price = 0
        self.buy_max_price = 0

        self.bool_marketday = False
        self.bool_stockorder = False
        self.bool_stockorder_timer = False
        self.bool_marketday_end = False
        self.bool_threshold = False

        self.init_marketday = None
        self.init_stockorder_timer = None


    def init_per_day(self):

        self.bkk = BotKIKr(self.key, self.secret, self.account, self.mock)
        self.bdf = self.load_xlsx(FILE_URL_DATA_15M).set_index('date')
        self.b_l = self.bdf.columns.to_list()
        self.q_l = self.get_guant_code_list()
        self.r_l = list(set(self.get_balance_code_list()).difference(self.q_l))

        self.tot_evl_price = self.get_total_price()
        self.buy_max_price = self.tot_evl_price / 20
        self.init_marketday = self.bkk.fetch_marketday()

        self.line_message(f'Bot15Self \n평가금액 : {self.total_eval_price}원, 다른종목: {len(self.r_l)}개, 위험종목 : {len(self.C_LIST)}개')


    def get_total_price(self):
        _total_eval_price = int(self.bkk.fetch_balance()['output2'][0]['tot_evlu_amt'])
        return _total_eval_price if _total_eval_price < 20000000 else 20000000
    

    def stock_order(self):

        tn = datetime.datetime.now()
        tn_153000 = tn.replace(hour=15, minute=30, second=0)
        tn_div = tn.minute % 15
        tn_del = None

        if tn_div == 0:
            tn_del = 1
        elif tn_div == 1:
            tn_del = 2
        elif tn_div == 2:
            tn_del = 3
        elif tn_div == 3:
            tn_del = 4
        elif tn_div == 4:
            tn_del = 5
        elif tn_div == 5:
            tn_del = 6
        elif tn_div == 6:
            tn_del = 7
        elif tn_div == 7:
            tn_del = 8
        elif tn_div == 8:
            tn_del = 9
        elif tn_div == 9:
            tn_del = 10
        elif tn_div == 10:
            tn_del = 11
        elif tn_div == 11:
            tn_del = 12
        elif tn_div == 12:
            tn_del = 13
        elif tn_div == 13:
            tn_del = 14
        elif tn_div == 14:
            tn_del = 15

        tn_del_min = tn - datetime.timedelta(minutes=tn_del)
        tn_df_idx = tn_del_min.strftime('%Y%m%d%H%M00') if tn < tn_153000 else tn.strftime('%Y%m%d153000')
        tn_df_req = tn_del_min.strftime('%H%M00') if tn < tn_153000 else '153000'

        print('##################################################')

        if self.bool_threshold and tn_div == 14:
            self.bdf = self.bdf[:-1]
        self.bool_threshold = False

        bal_lst = self.get_balance_code_list(True)
        sel_lst = []

        for code in self.b_l:

            min_lst = self.bkk.fetch_today_1m_ohlcv(code, tn_df_req, True)['output2'][:15]
            cur_prc = min_lst[0]['stck_prpr']
            self.bdf.at[tn_df_idx, code] = cur_prc
            
            is_late = tn_div == 2 or tn_div == 3 or tn_div == 4 or tn_div == 5 or tn_div == 6 or tn_div == 7 or tn_div == 8 or tn_div == 9 or tn_div == 10 or tn_div == 11 or tn_div == 12 or tn_div == 13 or tn_div == 14

            if (not is_late):

                is_remain = code in self.r_l
                is_alread = code in bal_lst
                is_buy = self.can_i(code, self.bdf[code], 'buy')
                is_sel = self.can_i(code, self.bdf[code], 'sel')
                
                if is_buy and (not is_alread) and (not is_remain):
                    
                    ord_q = self.get_qty(int(cur_prc), self.buy_max_price)
                    buy_r = self.bkk.create_market_buy_order(code, ord_q) if tn < tn_153000 else self.bkk.create_over_buy_order(code, ord_q)

                    if buy_r['rt_cd'] == '0':
                        print(f'매수 - 종목: {code}, 수량: {ord_q}주')
                        sel_lst.append({'c': '[B] ' + code, 'r': str(ord_q) + '주'})
                    else:
                        msg = buy_r['msg1']
                        print(f'{msg}')

                if is_sel and is_alread:

                    sel_r = self.bkk.create_market_sell_order(code, bal_lst[code]['q']) if tn < tn_153000 else self.bkk.create_over_sell_order(code, bal_lst[code]['q'])
                    _ror = self.ror(bal_lst[code]['ptp'], bal_lst[code]['ctp'])

                    if sel_r['rt_cd'] == '0':
                        print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                        sel_lst.append({'c': '[S] ' + code, 'r': round(_ror, 4)})

                    else:
                        msg = sel_r['msg1']
                        print(f'{msg}')

        sel_txt = ''
        for sl in sel_lst:
            sel_txt = sel_txt + '\n' + str(sl['c']) + ' : ' + str(sl['r'])
        
        _tn = datetime.datetime.now()
        _tn_151500 = _tn.replace(hour=15, minute=15, second=0)
        _tn_div = _tn.minute % 15
        _tn_sec = _tn.second
        _tn_del = None

        if _tn_div == 0:
            if tn_div == 14: 
                _tn_del = 0
                _tn_sec = 0
            else:
                _tn_del = 15
        elif _tn_div == 1:
            if tn_div == 14: 
                _tn_del = 0
                _tn_sec = 0
            else:
                _tn_del = 14
        elif _tn_div == 2:
            _tn_del = 13
        elif _tn_div == 3:
            _tn_del = 12
        elif _tn_div == 4:
            _tn_del = 11
        elif _tn_div == 5:
            _tn_del = 10
        elif _tn_div == 6:
            _tn_del = 9
        elif _tn_div == 7:
            _tn_del = 8
        elif _tn_div == 8:
            _tn_del = 7
        elif _tn_div == 9:
            _tn_del = 6
        elif _tn_div == 10:
            _tn_del = 5
        elif _tn_div == 11:
            _tn_del = 4
        elif _tn_div == 12:
            _tn_del = 3
        elif _tn_div == 13:
            _tn_del = 2
        elif _tn_div == 14:
            _tn_del = 1

        if _tn > _tn_151500:
            self.init_stockorder_timer = threading.Timer((60 * (30 - _tn.minute)) - _tn_sec, self.stock_order)
        else:
            self.init_stockorder_timer = threading.Timer((60 * _tn_del) - _tn_sec, self.stock_order)

        if self.bool_stockorder_timer:
            self.init_stockorder_timer.cancel()

        self.init_stockorder_timer.start()

        self.line_message(f'Bot15Self \n시작 : {tn}, \n표기 : {tn_df_idx} \n종료 : {_tn}, {sel_txt}')

    
    def get_min_df(self, code, to, min):

        df = None
        a_d = []
        a_c = []

        min_lst = self.fetch_today_1m_ohlcv(code, to)['output2']

        for i, m in enumerate(min_lst):
            a_d.append(str(m['stck_bsop_date'] + m['stck_cntg_hour']))
            a_c.append((m['stck_prpr']))

        df = pd.DataFrame({'date': a_d, code: a_c})
        df = df.set_index('date')

        n_s = 0
        if min == 3:
            n_s = 10
        elif min == 5 or min == 10:
            n_s = 11
        elif min == 15:
            n_s = 16
            
        if to == '153000':
            df_h = df.head(1)
            df_b = df.iloc[n_s::min, :]
            df = pd.concat([df_h, df_b])[::-1]
        else:
            df = df.iloc[::min, :][::-1]

        return df


    def market_to_excel(self, rebalance=False):

        _code_list = list(set(self.get_guant_code_list() + self.get_balance_code_list()))

        tn = datetime.datetime.now()
        if rebalance:
            tn = tn.replace(hour=15, minute=30, second=0)
        tn_093000 = tn.replace(hour=9, minute=30, second=0)
        
        if tn > tn_093000:

            tn_div = tn.minute % 15
            tn_del = None

            if tn_div == 0:
                tn_del = 16
            elif tn_div == 1:
                tn_del = 17
            elif tn_div == 2:
                tn_del = 18
            elif tn_div == 3:
                tn_del = 19
            elif tn_div == 4:
                tn_del = 20
            elif tn_div == 5:
                tn_del = 21
            elif tn_div == 6:
                tn_del = 22
            elif tn_div == 7:
                tn_del = 23
            elif tn_div == 8:
                tn_del = 24
            elif tn_div == 9:
                tn_del = 25
            elif tn_div == 10:
                tn_del = 26
            elif tn_div == 11:
                tn_del = 27
            elif tn_div == 12:
                tn_del = 28
            elif tn_div == 13:
                tn_del = 29
            elif tn_div == 14:
                tn_del = 15

            tn_req = ''
            tn_int = int(tn.strftime('%H%M%S'))
            tn_pos_a = 153000 <= tn_int
            tn_pos_b = 151500 < tn_int and tn_int < 153000
            tn_pos_c = tn_int <= 151500

            if tn_pos_a:
                tn_req = '153000'
            elif tn_pos_b:
                tn_req = '151400'
            elif tn_pos_c:
                tn_req = (tn - datetime.timedelta(minutes=tn_del)).strftime('%H%M00')
            
            df_a = []
            for c, code in enumerate(_code_list):
                print(f"{c + 1}/{len(_code_list)} {code}")
                df_a.append(self.get_min_df(code, tn_req, 15))
            df = pd.concat(df_a, axis=1)
            df = df.loc[~df.index.duplicated(keep='last')]

            print('##################################################')
            self.line_message(f'File Download Complete : {FILE_URL_DATA_15M}')
            print(df)
            df.to_excel(FILE_URL_DATA_15M)

            _tn = datetime.datetime.now()
            _tn_div = _tn.minute % 15

            if tn_pos_c and _tn_div == 14:
                self.bool_threshold = True

    
    def deadline_to_excel(self):
        kp = self.bkk.fetch_kospi_symbols()
        kp = kp.loc[(kp['그룹코드'] == 'ST') 
                    & (kp['시가총액규모'] != 0) 
                    & (kp['시가총액'] != 0) 
                    & (kp['우선주'] == 0) 
                    & (kp['단기과열'] == 0) 
                    & (kp['락구분'] == 0)
                    & (kp['액면변경'] == 0) 
                    & (kp['증자구분'] == 0)
                    & (kp['ETP'] != 'Y') 
                    & (kp['SRI'] != 'Y') 
                    & (kp['ELW발행'] != 'Y') 
                    & (kp['KRX은행'] != 'Y') 
                    & (kp['KRX증권'] != 'Y')
                    & (kp['KRX섹터_보험'] != 'Y')
                    & (kp['SPAC'] != 'Y') 
                    & (kp['거래정지'] != 'Y') 
                    & (kp['정리매매'] != 'Y') 
                    & (kp['관리종목'] != 'Y') 
                    & (kp['시장경고'] != 'Y') 
                    & (kp['경고예고'] != 'Y') 
                    & (kp['불성실공시'] != 'Y') 
                    & (kp['우회상장'] != 'Y') 
                    & (kp['공매도과열'] != 'Y') 
                    & (kp['이상급등'] != 'Y') 
                    & (kp['회사신용한도초과'] != 'Y') 
                    & (kp['담보대출가능'] != 'Y') 
                    & (kp['대주가능'] != 'Y') 
                    & (kp['신용가능'] == 'Y')
                    & (kp['증거금비율'] != 100)
                    & (kp['전일거래량'] > 300000) 
                    ]
        kd = self.bkk.fetch_kosdaq_symbols()
        kd = kd.loc[(kd['그룹코드'] == 'ST') 
                    & (kd['시가총액규모'] != 0) 
                    & (kd['시가총액'] != 0) 
                    & (kd['우선주'] == 0) 
                    & (kd['단기과열'] == 0) 
                    & (kd['락구분'] == 0)
                    & (kd['액면변경'] == 0) 
                    & (kd['증자구분'] == 0)
                    & (kd['ETP'] != 'Y') 
                    & (kd['KRX은행'] != 'Y') 
                    & (kd['KRX증권'] != 'Y')
                    & (kd['KRX섹터_보험'] != 'Y')
                    & (kd['SPAC'] != 'Y') 
                    & (kd['투자주의'] != 'Y') 
                    & (kd['거래정지'] != 'Y') 
                    & (kd['정리매매'] != 'Y') 
                    & (kd['관리종목'] != 'Y') 
                    & (kd['시장경고'] != 'Y') 
                    & (kd['경고예고'] != 'Y') 
                    & (kd['불성실공시'] != 'Y') 
                    & (kd['우회상장'] != 'Y') 
                    & (kd['공매도과열'] != 'Y') 
                    & (kd['이상급등'] != 'Y') 
                    & (kd['회사신용한도초과'] != 'Y') 
                    & (kd['담보대출가능'] != 'Y') 
                    & (kd['대주가능'] != 'Y') 
                    & (kd['신용가능'] == 'Y')
                    & (kd['증거금비율'] != 100)
                    & (kd['전일거래량'] > 300000)
                    ]
        _code_list = kp['단축코드'].to_list() + kd['단축코드'].to_list()
        code_list = self.get_caution_code_list(_code_list, True)

        self.save_file(FILE_URL_QUANT_LAST_15M, code_list)
        self.market_to_excel(True)


    def save_xlsx(self, url, df):
        df.to_excel(url)


    def load_xlsx(self, url):
        return pd.read_excel(url)
    

    def save_file(self, url, obj):
        with open(url, 'wb') as f:
            pickle.dump(obj, f)

    
    def load_file(self, url):
        with open(url, 'rb') as f:
            return pickle.load(f)
        

    def delete_file(self, url):
        if os.path.exists(url):
            for file in os.scandir(url):
                os.remove(file.path)

    
    def get_qty(self, crnt_p, max_p):
        q = int(max_p / int(crnt_p))
        return 1 if q == 0 else q
        
    
    def get_balance_code_list(self, obj=False):
        l = self.bkk.fetch_balance()['output1']
        a = []
        o = {}
        if len(l) > 0:
            for i in l:
                if int(i['ord_psbl_qty']) != 0:
                    if obj:
                        p = i['prpr']
                        q = i['ord_psbl_qty']
                        a = i['pchs_avg_pric']
                        o[i['pdno']] = {
                            'q': q,
                            'p': p,
                            'a': a,
                            'ptp': float(a) * int(q),
                            'ctp': float(p) * int(q)
                        }
                    else:
                        a.append(i['pdno'])
        return o if obj else a
    
    
    def get_guant_code_list(self):
        _l = self.load_file(FILE_URL_QUANT_LAST_15M)
        l = [str(int(i)).zfill(6) for i in _l]
        return l
    

    def get_caution_code_list(self, l, rm=False):
        a = []
        for _l in l:
            r = self.bkk.fetch_price(_l)['output']['iscd_stat_cls_code']
            if (r == '51') or (r == '52') or (r == '53') or (r == '54') or (r == '58') or (r == '59'):
                if rm:
                    l.remove(_l)
                else:
                    a.append(_l)

        return l if rm else a
    

    def rsi(self, df, period=14):
        _f = df.head(1)
        _o = {}
        dt = df.diff(1).dropna()
        u, d = dt.copy(), dt.copy()
        u[u < 0] = 0
        d[d > 0] = 0
        _o['u'] = u
        _o['d'] = d
        au = _o['u'].rolling(window = period).mean()
        ad = abs(_o['d'].rolling(window = period).mean())
        rs = au / ad
        _rsi = pd.Series(100 - (100 / (1 + rs)))
        rsi = pd.concat([_f, _rsi])
        return rsi


    def rsi_vol_zremove(self, df, code):
        _a = []
        for i, d in df.iterrows():
            if d[code].split('|')[1] != '0':
                _a.append(int(d[code].split('|')[0]))
        df_c = pd.DataFrame({'close': _a})
        _rsi = self.rsi(df_c['close']).iloc[-1]
        rsi = 'less' if math.isnan(_rsi) else _rsi
        return rsi
    

    def ror(self, pv, nv, pr=1, pf=0.00015, spf=0.003):
        cr = ((nv - (nv * pf) - (nv * spf)) / (pv + (pv * pf)))
        return pr * cr
    
    
    def line_message(self, msg):
        print(msg)
        requests.post(LINE_URL, headers={'Authorization': 'Bearer ' + LINE_TOKEN}, data={'message': msg})
    

if __name__ == '__main__':

    B15 = Bot_15m()
    # B15.market_to_excel()
    # B15.deadline_to_excel()

    while True:

        try:

            t_n = datetime.datetime.now()
            t_085000 = t_n.replace(hour=8, minute=50, second=0)
            t_091500 = t_n.replace(hour=9, minute=15, second=0)
            t_152500 = t_n.replace(hour=15, minute=25, second=0)
            t_153000 = t_n.replace(hour=15, minute=30, second=0)
            t_160000 = t_n.replace(hour=16, minute=0, second=0)

            if t_n >= t_085000 and t_n <= t_153000 and B15.bool_marketday == False:
                if os.path.isfile(os.getcwd() + '/token.dat'):
                    os.remove('token.dat')
                B15.init_per_day()
                B15.bool_marketday = True
                B15.bool_marketday_end = False

                B15.line_message(f'Stock Start' if B15.init_marketday == 'Y' else 'Holiday Start')

            if B15.init_marketday == 'Y':

                if t_n > t_152500 and t_n < t_153000 and B15.bool_stockorder_timer == False:
                    B15.bool_stockorder_timer = True

                if t_n >= t_091500 and t_n <= t_153000 and B15.bool_stockorder == False:
                    B15.stock_order()
                    B15.bool_stockorder = True

            if t_n == t_160000 and B15.bool_marketday_end == False:

                if B15.init_marketday == 'Y':
                    B15.deadline_to_excel()
                    B15.bool_stockorder_timer = False
                    B15.bool_stockorder = False

                B15.bool_marketday = False
                B15.bool_marketday_end = True

                B15.line_message(f'Stock End' if B15.init_marketday == 'Y' else 'Holiday End')

        except Exception as e:

            B15.line_message(f"Bot15 Error : {e}")
            break