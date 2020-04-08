#########################################
# Author: https://github.com/Vladqwelkj #
# Telegram: @Vladqwelkj                 #
#########################################


import time
import os
import datetime
import json
import threading

import requests
import websocket
from binance.client import Client
import openpyxl
from pathlib import Path

SYMBOL = str()


def in_new_thread(my_func):
    def wrapper(*args, **kwargs):
        my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
        my_thread.start()
    return wrapper


def write_log(*args):
    input_data = ' '.join([str(v) for v in args])
    print(input_data)
    open('log.log', 'a').write(datetime.datetime.strftime(datetime.datetime.now(), '[%d/%m/%Y %H:%M:%S] ') 
        + input_data + '\n') #Текущая дата/время + текст


class XlsxParser:
    def __init__(self, filename):
        global write_log
        try:
            write_log('Start .xlsx parsing')
            xlsx_file = Path('settings.xlsx')
            wb_obj = openpyxl.load_workbook(xlsx_file) 
            sheet = wb_obj.active

            self.api_key = str(sheet['B1'].value)
            self.api_secret = str(sheet['B2'].value)
            self.symbol = str(sheet['B3'].value)
            self.settings_list = []
            for ind, row in enumerate(sheet):
                if ind < 4:
                    continue
                for cell in row:
                    if cell.value==None:
                        write_log('Ошибка! В xlsx файле пустая ячейка на '+str(cell))
                        exit()
                self.settings_list.append(type('SettingsSet', (), {
                    'level': float(row[0].value),
                    'amount': float(row[1].value),
                    'tp_level': float(row[2].value),
                    'trail': float(row[3].value),
                    #'symbol': row[4].value,
                    }))
            write_log('.xlsx parsed')
        except Exception as e:
            write_log('parsing .xlsx ERROR!:', e)
    def get_result(self):

        return type('ParamSet', (), {
            'api_key': self.api_key,
            'api_secret': self.api_secret,
            'symbol': self.symbol,
            'trading_settings': self.settings_list,})


def test_connection():
    global write_log
    write_log('Тест соединения к бирже')
    response = requests.get('https://api.binance.com/')
    if response.status_code==200:
        write_log('Тест соединения пройден')
        return True
    write_log('Соединение не удалось')
    exit()



class Order:
    def __init__(self, client, symbol, side, price, amount, ordertype):
        '''Support only LIMIT and STOP_MARKET orders. Supports price editting'''
        if ordertype != 'LIMIT' and ordertype != 'STOP_MARKET':
            write_log('Class Order: Supports only LIMIT and STOP_MARKETS orders!')
            exit()

        self.symbol = symbol
        self.ordertype = ordertype
        self.side = side
        self.amount = amount

        limit_price = None if ordertype=='STOP_MARKET' else str(price)
        stop_price = str(price) if 'STOP_MARKET'==ordertype else None
        self._price = limit_price if stop_price==None else stop_price
        self.client = client
        if ordertype=='LIMIT':
            self.order_id = client.create_order(
                symbol=symbol,
                side=side,
                type=ordertype,
                quantity=amount,
                price=limit_price,
                timeInForce='GTC')['orderId']
        if ordertype=='STOP_MARKET':
            self.order_id = client.create_order(
               symbol=symbol,
               side=side,
               type='STOP_LOSS_LIMIT',
               quantity=amount,
               price=str(float(stop_price)*1.01),
               stopPrice=stop_price,
               timeInForce='GTC')['orderId']
        write_log('created', symbol, side, price, amount, ordertype, self.order_id)


    @property
    def status(self):
        #write_log('req for status, id', self.order_id)
        return self.client.get_order(symbol=self.symbol, orderId=self.order_id)['status']
    
    @property
    def price(self):
        return float(self._price)

    @price.setter
    def price(self, new_price):
        try:
            self.client.cancel_order(symbol=self.symbol, orderId=self.order_id)
            limit_price = str(new_price) if self.ordertype=='LIMIT' else None
            stop_price = str(new_price) if self.ordertype=='STOP_MARKET' else None
            self._price = str(new_price)

            if ordertype=='LIMIT':
                self.order_id = client.create_order(
                    symbol=self.symbol,
                    side=self.side,
                    type=self.ordertype,
                    quantity=self.amount,
                    price=limit_price,
                    timeInForce='GTC')['orderId']
            if ordertype=='STOP_MARKET':
                self.order_id = client.create_order(
                   symbol=self.self.symbol,
                   side=self.side,
                   type='STOP_LOSS_LIMIT',
                   quantity=self.amount,
                   price=str(float(stop_price)*1.01),
                   stopPrice=stop_price,
                   timeInForce='GTC')['orderId']

            write_log(self.order_id, 'price edited')
            return self.order_id
        except Exception as e:
            print('price_setter ERROR:', e)

    def __del__(self):
        try:
            self.client.cancel_order(symbol=self.symbol, orderId=self.order_id)
            write_log(self.order_id, 'canceled')
        except Exception as e:
            pass
        #write_log('delete(cancel) order error:', e)



class TrailingOrdersManager:
    def __init__(self, client, orders_caretaker):
        global SYMBOL
        self.client = client
        self.orders_caretaker = orders_caretaker
        self.levels_and_trail_orders = {} # {Level: trail_order, Level2: trail_order2, ...}

    def _price_now(self):
        price = self.client.get_klines(
            symbol=SYMBOL, interval='1m', limit=1)[-1][2]
        return float(price)

    @in_new_thread
    def run(self):
        global SYMBOL
        while True:
            time.sleep(5)
            price = self._price_now()
            for level in self.levels_and_trail_orders.keys():
                trail_order = self.levels_and_trail_orders[level]
                if price > level.level+level.tp_level and not trail_order:
                    write_log(level.level, ': place trail order')
                    self.levels_and_trail_orders[level] = Order(
                        client=self.client,
                        symbol=SYMBOL,
                        side='SELL',
                        price=level.level - level.trail,
                        amount=level.amount,
                        ordertype='STOP_MARKET')

                    self.orders_caretaker.orders_list.append({
                        'order': self.levels_and_trail_orders[level],
                        'func_for_filled_order': level.level_order_setup})
                if trail_order and price > trail_order.price:
                    write_log(level.level, ': trail order taut')
                    self.levels_and_trail_orders[level].price = price
            



class OrdersCaretaker:
    def __init__(self, client):
        self.client = client
        self.orders_list = [] #{'order':Order,'func_for_filled_order':func}, ...]

    @in_new_thread
    def run(self):
        global SYMBOL
        while True:
            orders = client.get_all_orders(symbol=SYMBOL, limit=200)

            for order in orders:
                for level in self.orders_list:
                    if (order['orderId']==level['order'].order_id
                        and order['status']=='FILLED'):

                        write_log(order['orderId'], 'filled')
                        level['func_for_filled_order']()
                        self.orders_list.remove(level)
                        break
            time.sleep(20)



class LevelManager:
    def __init__(self, client, trailing_orders_manager, orders_caretaker, price_now, level, amount, tp_level, trail):
        global write_log, SYMBOL
        self.client = client
        self.trailing_orders_manager = trailing_orders_manager
        self.orders_caretaker = orders_caretaker
        self.price_now = price_now
        self.level = level
        self.amount = amount
        self.tp_level = tp_level
        self.trail = trail
        self.trailing_order = None

    @in_new_thread
    def start_working(self, ):
        write_log(str(self.level), 'start')
        self.level_order_setup(not_first=False)

    def _price_now(self):
        price = self.client.get_klines(
            symbol=SYMBOL, interval='1m', limit=1)[-1][2]
        return float(price)

    def level_order_setup(self, not_first=True):
        global SYMBOL
        if not_first:
            price_now = self._price_now()
        if self.price_now <= self.level:
            order = Order(
                client=self.client,
                symbol=SYMBOL,
                side='BUY',
                price=self.level,
                amount=self.amount,
                ordertype='STOP_MARKET')
        else:
            order = Order(
                client=self.client,
                symbol=SYMBOL,
                side='BUY',
                price=self.level,
                amount=self.amount,
                ordertype='LIMIT')
        write_log(str(self.level), ': order placed, id:', order.order_id)
        self.orders_caretaker.orders_list.append({
            'order': order,
            'func_for_filled_order': self.do_when_level_order_filled})


    def do_when_level_order_filled(self):
        self.trailing_orders_manager.levels_and_trail_orders[self] = False


def cancel_all_orders(client, symbol):
	write_log('cancel all open orders')
	orders = client.get_all_orders(symbol=symbol, limit=200)
	for order in orders:
		if order['status']=='NEW':
			client.cancel_order(symbol=symbol, orderId=order['orderId'])


if __name__=='__main__':
    write_log('\n--------\nStart program')
    test_connection()

    params = XlsxParser(filename='settings.xlsx').get_result()
    
    SYMBOL = params.symbol
    write_log(SYMBOL)
    client = Client(params.api_key, params.api_secret)
    cancel_all_orders(client, symbol=SYMBOL)
    orders_caretaker = OrdersCaretaker(client)
    trailing_orders_manager = TrailingOrdersManager(client, orders_caretaker)
    price_now = float(client.get_klines(
        symbol=SYMBOL, interval='1m', limit=1)[-1][2])
    for p in params.trading_settings:
        level_manager = LevelManager(
            client=client,
            trailing_orders_manager=trailing_orders_manager,
            orders_caretaker=orders_caretaker,
            price_now=price_now,
            level=p.level,
            amount=p.amount,
            tp_level=p.tp_level,
            trail=p.trail,)
        level_manager.start_working()
    trailing_orders_manager.run()
    orders_caretaker.run()
