    # def market_to_excel(self, rebalance=False, filter=False):

    #     tn = datetime.datetime.now()
    #     if rebalance:
    #         tn = tn.replace(hour=15, minute=30, second=0)
    #     tn_093000 = tn.replace(hour=9, minute=30, second=0)
        
    #     if tn > tn_093000:

    #         tn_div = tn.minute % 15
    #         tn_del = None

    #         if tn_div == 0:
    #             tn_del = 16
    #         elif tn_div == 1:
    #             tn_del = 17
    #         elif tn_div == 2:
    #             tn_del = 18
    #         elif tn_div == 3:
    #             tn_del = 19
    #         elif tn_div == 4:
    #             tn_del = 20
    #         elif tn_div == 5:
    #             tn_del = 21
    #         elif tn_div == 6:
    #             tn_del = 22
    #         elif tn_div == 7:
    #             tn_del = 23
    #         elif tn_div == 8:
    #             tn_del = 24
    #         elif tn_div == 9:
    #             tn_del = 25
    #         elif tn_div == 10:
    #             tn_del = 26
    #         elif tn_div == 11:
    #             tn_del = 27
    #         elif tn_div == 12:
    #             tn_del = 28
    #         elif tn_div == 13:
    #             tn_del = 29
    #         elif tn_div == 14:
    #             tn_del = 15

    #         tn_req = ''
    #         tn_int = int(tn.strftime('%H%M%S'))
    #         tn_pos_a = 153000 <= tn_int
    #         tn_pos_b = 151500 < tn_int and tn_int < 153000
    #         tn_pos_c = tn_int <= 151500

    #         if tn_pos_a:
    #             tn_req = '153000'
    #         elif tn_pos_b:
    #             tn_req = '151400'
    #         elif tn_pos_c:
    #             tn_req = (tn - datetime.timedelta(minutes=tn_del)).strftime('%H%M00')

    #         if filter:
    #             fltr_list = self.bkk.filter_code_list()
    #             if len(fltr_list) > 0:
    #                 save_file(FILE_URL_SMBL_15M, fltr_list)

    #         _code_list = list(set(self.get_guant_code_list() + self.get_balance_code_list()))
            
    #         df_a = []
    #         for c, code in enumerate(_code_list):
    #             print(f"{c + 1}/{len(_code_list)} {code}")
    #             df_a.append(self.bkk.df_today_1m_ohlcv(code, tn_req, 15))
    #         df = pd.concat(df_a, axis=1)
    #         df = df.loc[~df.index.duplicated(keep='last')]

    #         print('##################################################')
    #         line_message(f'Bot15Swing Total Symbol Data: {len(_code_list)}개, \n{_code_list} \nFile Download Complete : {FILE_URL_DATA_15M}')
    #         print(df)
    #         df.to_excel(FILE_URL_DATA_15M)

    #         _tn = datetime.datetime.now()
    #         _tn_div = _tn.minute % 15

    #         if tn_pos_c and _tn_div == 14:
    #             self.bool_threshold = True

    
    # def deadline_to_excel(self):
    #     sym_lst = self.bkk.filter_code_list()
    #     if len(sym_lst) > 0:
    #         print('##################################################')
    #         line_message(f'Bot3Swing Symbol List: {len(sym_lst)}개, \n{sym_lst} \nFile Download Complete : {FILE_URL_SMBL_15M}')
    #         save_file(FILE_URL_SMBL_15M, sym_lst)