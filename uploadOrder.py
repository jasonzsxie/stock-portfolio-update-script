from ibapi.client import EClient
from ibapi.wrapper import EWrapper, OrderState
from ibapi.contract import Contract
from ibapi.order import Order
from ibapi.execution import Execution
import threading
from pathlib import Path
import time
import pandas as pd

class IBApi(EWrapper, EClient):
    def __init__(self):
        EClient.__init__(self, self)
        self.nextOrderId = None
        self.orders = []  # To keep track of placed orders
        self.contract_details = {}
        self.resolve_contracts_event = threading.Event()

    def nextValidId(self, orderId: int):
        self.nextOrderId = orderId
        self.start()

    def error(self, reqId, errorCode, errorString):
        print(f"Error {reqId} {errorCode} {errorString}")

    def orderStatus(self, orderId: int, status: str, filled: float,
                    remaining: float, avgFillPrice: float, permId: int,
                    parentId: int, lastFillPrice: float, clientId: int,
                    whyHeld: str, mktCapPrice: float):
        print(f"OrderStatus. Id: {orderId}, Status: {status}, Filled: {filled}, Remaining: {remaining}, LastFillPrice: {lastFillPrice}")

    def openOrder(self, orderId: int, contract: Contract, order: Order,
                  orderState: OrderState):
        print(f"OpenOrder. Id: {orderId}, {contract.symbol}, {contract.secType}, {order.action}, {order.orderType}, {order.totalQuantity}, {order.lmtPrice}")

    def execDetails(self, reqId: int, contract: Contract, execution: Execution):
        print(f"ExecDetails. {reqId}, {contract.symbol}, {contract.secType}, {execution.execId}, {execution.orderId}, {execution.shares}, {execution.lastLiquidity}")

    def connectAck(self):
        if self.asynchronous:
            self.startApi()

    def contractDetails(self, reqId: int, contractDetails):
        print(f"Received contract details for request {reqId}: {contractDetails}")
        if reqId not in self.contract_details:
            self.contract_details[reqId] = []
        self.contract_details[reqId].append(contractDetails)

    def contractDetailsEnd(self, reqId: int):
        print(f"Contract details end for request {reqId}")
        self.resolve_contracts_event.set()

    def resolve_contract(self, contract):
        self.resolve_contracts_event.clear()
        reqId = self.nextOrderId
        self.contract_details[reqId] = []
        print(f"Requesting contract details for {contract.symbol}")
        self.reqContractDetails(reqId, contract)
        self.resolve_contracts_event.wait(timeout=10)
        if reqId in self.contract_details:
            return self.contract_details[reqId]
        else:
            print(f"No contract details returned for {contract.symbol}")
            return []
    
    def start(self):
        ROOT_DIR = Path(__file__).parent
        pathFile = ROOT_DIR / 'updated_portfolio.xlsx'
        df = pd.read_excel(pathFile, header = 0)

        for index, row in df.iterrows():
            ticker = row['Ticker']
            closing_price = row['Closing Price']
            shares = row['Shares to Buy/Sell']

            if shares == 0:
                print(f"No modifification is needed for stock {ticker}")
                continue

            # Define the contract
            contract = Contract()
            contract.symbol = ticker
            contract.secType = "STK"
            contract.exchange = "SMART"
            contract.currency = "USD"

            # Resolve contract ambiguity
            #resolved_contract = self.resolve_contract(contract)
            #if not resolved_contract or len(resolved_contract) == 0:
           #     print(f"Could not resolve contract for {ticker}. Skipping.")
           #     continue

            # Select the contract to use
            #resolved_contract = resolved_contract[0].contract
           # if len(resolved_contract) > 1:
           #     for details in resolved_contract:
           #         if details.contract.primaryExchange == "NASDAQ":
            #            resolved_contract = details.contract
            #            break

            # Define the order
            order = Order()
            order.orderType = "LMT"
            order.totalQuantity = abs(shares)
            
            if shares > 0:
                order.action = "BUY"
                order.lmtPrice = round(closing_price * 0.8, 2)
            elif shares < 0:
                order.action = "SELL"
                order.lmtPrice = round(closing_price * 1.2, 2)

            order.tif = "GTC"

            order.eTradeOnly = False
            order.firmQuoteOnly = False
            print(f"Placing order for {ticker}: {order.action} {order.totalQuantity} shares at {order.lmtPrice}")

            self.placeOrder(self.nextOrderId, contract, order)
            self.nextOrderId += 1
            time.sleep(1)  # Sleep for 1 second between orders to ensure they're processed

        

    def stop(self):
        self.done = True
        self.disconnect()

def run_loop():
    app.run()

app = IBApi()
app.connect("127.0.0.1", 7497, 0)  # Use port 4002 for paper trading, USE PORT 7497 for live trading

api_thread = threading.Thread(target=run_loop, daemon=True)
api_thread.start()

time.sleep(60)
app.stop()
