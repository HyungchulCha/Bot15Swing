from BotConfig import *
from BotUtil import *
from BotKIKr import BotKIKr
from dateutil.relativedelta import *
import pandas as pd
import datetime
import threading
import os
import copy

class Bot15Swing():

    
    def __init__(self):

        self.mock = False
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
        self.bdf = load_xlsx(FILE_URL_DATA_15M).set_index('date')
        self.b_l = self.bdf.columns.to_list()
        self.q_l = self.get_guant_code_list()
        self.r_l = list(set(self.get_balance_code_list()).difference(self.q_l))

        _ttl_prc = int(self.bkk.fetch_balance()['output2'][0]['tot_evlu_amt'])
        _buy_cnt = len(self.q_l) if len(self.q_l) > 30 else 30
        _buy_cnt = 60
        
        self.tot_evl_price = _ttl_prc if _ttl_prc < 30000000 else 30000000
        self.buy_max_price = self.tot_evl_price / _buy_cnt
        self.init_marketday = self.bkk.fetch_marketday()

        line_message(f'Bot15Swing \n평가금액 : {self.tot_evl_price}원, 다른종목: {len(self.r_l)}개')
    

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
        # tn_df_idx = tn_del_min.strftime('%Y%m%d%H%M00') if tn < tn_153000 else tn.strftime('%Y%m%d153000')
        # tn_df_req = tn_del_min.strftime('%H%M00') if tn < tn_153000 else '153000'
        tn_df_idx = tn_del_min.strftime('%Y%m%d%H%M00')
        tn_df_req = tn_del_min.strftime('%H%M00')

        print('##################################################')

        if self.bool_threshold and tn_div == 14:
            self.bdf = self.bdf[:-1]
        self.bool_threshold = False

        bal_lst = self.get_balance_code_list(True)
        sel_lst = []

        if os.path.isfile(FILE_URL_BLNC_15M):
            obj_lst = load_file(FILE_URL_BLNC_15M)
        else:
            obj_lst = {}
            save_file(FILE_URL_BLNC_15M, obj_lst)

        for code in self.b_l:

            min_lst = self.bkk.fetch_today_1m_ohlcv(code, tn_df_req, True)['output2'][:15]
            chk_cls = float(min_lst[0]['stck_prpr'])
            chk_opn = float(min_lst[14]['stck_oprc'])
            chk_hig = max([float(min_lst[i]['stck_hgpr']) for i in range(15)])
            chk_low = min([float(min_lst[i]['stck_lwpr']) for i in range(15)])
            chk_vol = sum([int(min_lst[i]['cntg_vol']) for i in range(15)])
            self.bdf.at[tn_df_idx, code] = str(chk_opn) + '|' + str(chk_hig) + '|' + str(chk_low) + '|' + str(chk_cls) + '|' + str(chk_vol)
            
            is_late = tn_div == 2 or tn_div == 3 or tn_div == 4 or tn_div == 5 or tn_div == 6 or tn_div == 7 or tn_div == 8 or tn_div == 9 or tn_div == 10 or tn_div == 11 or tn_div == 12 or tn_div == 13 or tn_div == 14

            if (not is_late):

                is_remain = code in self.r_l
                is_alread = code in bal_lst

                if is_alread and not (code in obj_lst):
                    obj_lst[code] = {'x': copy.deepcopy(bal_lst[code]['a']), 'a': copy.deepcopy(bal_lst[code]['a']), 's': 1, 'd': datetime.datetime.now().strftime('%Y%m%d')}

                if not is_alread and (code in obj_lst):
                    obj_lst.pop(code, None)
                
                if (not is_alread) and (not is_remain):

                    df = min_max_height(moving_average(get_code_df(self.bdf, code)))
                    df_t = df.tail(1)

                    if \
                    (df_t['close'].iloc[-1] < (df_t['close_p'].iloc[-1] * 1.05)) and \
                    (df_t['height'].iloc[-1] > 1.1) and \
                    (df_t['ma05'].iloc[-1] > df_t['ma20'].iloc[-1] > df_t['ma60'].iloc[-1]) and \
                    (df_t['ma20'].iloc[-1] * 1.05 > df_t['close'].iloc[-1] > df_t['ma20'].iloc[-1]) and \
                    (df_t['close'].iloc[-1] > df_t['ma05'].iloc[-1])\
                    :
                        if chk_cls < self.buy_max_price:

                            ord_q = get_qty(chk_cls, self.buy_max_price)
                            buy_r = self.bkk.create_market_buy_order(code, ord_q) if tn < tn_153000 else self.bkk.create_over_buy_order(code, ord_q)

                            if buy_r['rt_cd'] == '0':
                                print(f'매수 - 종목: {code}, 수량: {ord_q}주')
                                obj_lst[code] = {'a': chk_cls, 'x': chk_cls, 's': 1, 'd': datetime.datetime.now().strftime('%Y%m%d')}
                                sel_lst.append({'c': '[B] ' + code, 'r': str(ord_q) + '주'})
                            else:
                                msg = buy_r['msg1']
                                print(f'{msg}')

                obj_ntnul = not (not obj_lst)

                if is_alread and obj_ntnul:

                    obj_d = obj_lst[code]['d']
                    now_d = datetime.datetime.now().strftime('%Y%m%d')
                    dif_d = datetime.datetime(int(now_d[:4]), int(now_d[4:6]), int(now_d[6:])) - datetime.datetime(int(obj_d[:4]), int(obj_d[4:6]), int(obj_d[6:]))

                    if (dif_d.days) >= 14:

                        bal_fst = bal_lst[code]['a']
                        bal_cur = bal_lst[code]['p']
                        bal_qty = bal_lst[code]['q']

                        sel_r = self.bkk.create_market_sell_order(code, bal_qty) if tn < tn_153000 else self.bkk.create_over_sell_order(code, bal_qty)
                        _ror = ror(bal_fst * bal_qty, bal_cur * bal_qty)

                        if sel_r['rt_cd'] == '0':
                            print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                            sel_lst.append({'c': '[SL] ' + code, 'r': round(_ror, 4)})
                            obj_lst.pop(code, None)
                        else:
                            msg = sel_r['msg1']
                            print(f'{msg}')

                    else:

                        t1 = 0.06
                        t2 = 0.08
                        t3 = 0.1
                        ct = 0.8
                        hp = 100

                        if obj_lst[code]['x'] < bal_lst[code]['p']:
                            obj_lst[code]['x'] = copy.deepcopy(bal_lst[code]['p'])
                            obj_lst[code]['a'] = copy.deepcopy(bal_lst[code]['a'])

                        if obj_lst[code]['x'] > bal_lst[code]['p']:

                            bal_pft = bal_lst[code]['pft']
                            bal_fst = bal_lst[code]['a']
                            bal_cur = bal_lst[code]['p']
                            bal_qty = bal_lst[code]['q']
                            rto_01 = 0.2
                            rto_02 = (3/8)
                            ord_qty_01 = int(bal_qty * rto_01) if int(bal_qty * rto_01) != 0 else 1
                            ord_qty_02 = int(bal_qty * rto_02) if int(bal_qty * rto_02) != 0 else 1
                            is_qty_01 = bal_qty == ord_qty_01
                            is_qty_02 = bal_qty == ord_qty_02
                            obj_max = obj_lst[code]['x']
                            obj_fst = obj_lst[code]['a']
                            obj_pft = obj_max / obj_fst
                            los_dif = obj_pft - bal_pft
                            sel_cnt = copy.deepcopy(obj_lst[code]['s'])

                            if 1 < bal_pft < hp:

                                if (sel_cnt == 1) and (t1 <= los_dif):

                                    sel_r = self.bkk.create_market_sell_order(code, ord_qty_01) if tn < tn_153000 else self.bkk.create_over_sell_order(code, ord_qty_01)
                                    _ror = ror(bal_fst * ord_qty_01, bal_cur * ord_qty_01)

                                    if sel_r['rt_cd'] == '0':
                                        print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                                        sel_lst.append({'c': '[S1] ' + code, 'r': round(_ror, 4)})
                                        obj_lst[code]['s'] = sel_cnt + 1
                                        obj_lst[code]['d'] = datetime.datetime.now().strftime('%Y%m%d')

                                        if is_qty_01:
                                            obj_lst.pop(code, None)
                                    else:
                                        msg = sel_r['msg1']
                                        print(f'{msg}')
                                
                                elif (sel_cnt == 2) and (t2 <= los_dif):

                                    sel_r = self.bkk.create_market_sell_order(code, ord_qty_02) if tn < tn_153000 else self.bkk.create_over_sell_order(code, ord_qty_02)
                                    _ror = ror(bal_fst * ord_qty_02, bal_cur * ord_qty_02)

                                    if sel_r['rt_cd'] == '0':
                                        print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                                        sel_lst.append({'c': '[S2] ' + code, 'r': round(_ror, 4)})
                                        obj_lst[code]['s'] = sel_cnt + 1
                                        obj_lst[code]['d'] = datetime.datetime.now().strftime('%Y%m%d')

                                        if is_qty_02:
                                            obj_lst.pop(code, None)
                                    else:
                                        msg = sel_r['msg1']
                                        print(f'{msg}')

                                elif (sel_cnt == 3) and (t3 <= los_dif):
                                        
                                    sel_r = self.bkk.create_market_sell_order(code, bal_qty) if tn < tn_153000 else self.bkk.create_over_sell_order(code, bal_qty)
                                    _ror = ror(bal_fst * bal_qty, bal_cur * bal_qty)

                                    if sel_r['rt_cd'] == '0':
                                        print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                                        sel_lst.append({'c': '[S3] ' + code, 'r': round(_ror, 4)})
                                        obj_lst[code]['s'] = sel_cnt + 1
                                        obj_lst[code]['d'] = datetime.datetime.now().strftime('%Y%m%d')
                                        obj_lst.pop(code, None)
                                    else:
                                        msg = sel_r['msg1']
                                        print(f'{msg}')

                            elif hp <= bal_pft:

                                sel_r = self.bkk.create_market_sell_order(code, bal_qty) if tn < tn_153000 else self.bkk.create_over_sell_order(code, bal_qty)
                                _ror = ror(bal_fst * bal_qty, bal_cur * bal_qty)

                                if sel_r['rt_cd'] == '0':
                                    print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                                    sel_lst.append({'c': '[S+] ' + code, 'r': round(_ror, 4)})
                                    obj_lst.pop(code, None)
                                else:
                                    msg = sel_r['msg1']
                                    print(f'{msg}')

                            elif bal_pft <= ct:

                                sel_r = self.bkk.create_market_sell_order(code, bal_qty) if tn < tn_153000 else self.bkk.create_over_sell_order(code, bal_qty)
                                _ror = ror(bal_fst * bal_qty, bal_cur * bal_qty)

                                if sel_r['rt_cd'] == '0':
                                    print(f'매도 - 종목: {code}, 수익: {round(_ror, 4)}')
                                    sel_lst.append({'c': '[S-] ' + code, 'r': round(_ror, 4)})
                                    obj_lst.pop(code, None)
                                else:
                                    msg = sel_r['msg1']
                                    print(f'{msg}')

        save_file(FILE_URL_BLNC_15M, obj_lst)

        sel_txt = ''
        for sl in sel_lst:
            sel_txt = sel_txt + '\n' + str(sl['c']) + ' : ' + str(sl['r'])
        
        _tn = datetime.datetime.now()
        # _tn_151500 = _tn.replace(hour=15, minute=15, second=0)
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

        # if _tn > _tn_151500:
        #     self.init_stockorder_timer = threading.Timer((60 * (30 - _tn.minute)) - _tn_sec, self.stock_order)
        # else:
        self.init_stockorder_timer = threading.Timer((60 * _tn_del) - _tn_sec, self.stock_order)

        if self.bool_stockorder_timer:
            self.init_stockorder_timer.cancel()

        self.init_stockorder_timer.start()

        line_message(f'Bot15Swing \n시작 : {tn}, \n표기 : {tn_df_idx} \n종료 : {_tn}, {sel_txt}')


    def market_to_excel(self):

        b_list = self.get_balance_code_list()
        q_list = self.bkk.get_condition_code_list()
        f_list = list(set(q_list + b_list))
        f_df = get_yfinance_df(f_list, 15)
        print('##################################################')
        save_file(FILE_URL_SMBL_15M, q_list)
        save_xlsx(FILE_URL_DATA_15M, f_df)
        line_message(f'Bot15Swing \nQuant List: {len(q_list)}종목 \nBalance List: {len(b_list)}종목 \nTotal List: {len(f_list)}개, \n{f_list} \nFile Download Complete : {FILE_URL_DATA_15M}')
        print(f_df)
        
    
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
                            'q': int(q),
                            'p': float(p),
                            'a': float(a),
                            'pft': float(p)/float(a),
                            'ptp': float(a) * int(q),
                            'ctp': float(p) * int(q)
                        }
                    else:
                        a.append(i['pdno'])
        return o if obj else a
    
    
    def get_guant_code_list(self):
        _l = load_file(FILE_URL_SMBL_15M)
        l = [str(int(i)).zfill(6) for i in _l]
        return l
    

if __name__ == '__main__':

    B15 = Bot15Swing()
    # B15.market_to_excel()

    while True:

        try:

            t_n = datetime.datetime.now()
            t_083000 = t_n.replace(hour=8, minute=30, second=0)
            t_091500 = t_n.replace(hour=9, minute=15, second=0)
            t_152500 = t_n.replace(hour=15, minute=25, second=0)
            t_153000 = t_n.replace(hour=15, minute=30, second=0)
            t_160000 = t_n.replace(hour=16, minute=0, second=0)

            if t_n >= t_083000 and t_n <= t_153000 and B15.bool_marketday == False:
                if os.path.isfile(os.getcwd() + '/token.dat'):
                    os.remove('token.dat')
                B15.init_per_day()
                if B15.init_marketday == 'Y':
                    B15.market_to_excel()
                B15.bool_marketday = True
                B15.bool_marketday_end = False

                line_message(f'Bot15Swing Stock Start' if B15.init_marketday == 'Y' else 'Bot15Swing Holiday Start')

            if B15.init_marketday == 'Y':

                if t_n > t_152500 and t_n < t_153000 and B15.bool_stockorder_timer == False:
                    B15.bool_stockorder_timer = True

                if t_n >= t_091500 and t_n <= t_153000 and B15.bool_stockorder == False:
                    B15.stock_order()
                    B15.bool_stockorder = True

            if t_n == t_160000 and B15.bool_marketday_end == False:

                if B15.init_marketday == 'Y':
                    B15.bool_stockorder_timer = False
                    B15.bool_stockorder = False

                B15.bool_marketday = False
                B15.bool_marketday_end = True

                line_message(f'Bot15Swing Stock End' if B15.init_marketday == 'Y' else 'Bot15Swing Holiday End')

        except Exception as e:

            line_message(f"Bot15Swing Error : {e}")
            break