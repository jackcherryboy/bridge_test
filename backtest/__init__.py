import pandas as pd
import os

# 设置全局回测参数
INITIAL_CAPITAL = 1000000  # 初始本金
POSITION_SIZE = 0.2  # 每次买入本金的20%
SLIPPAGE = 0.0001  # 滑点万分之一

# 文件路径设置
file_path = r'C:\Users\FUNDS\Desktop\金融研一作业\策略\000001.xlsx'


# 数据加载函数
def load_data(file_path):
    """
    加载Excel数据文件，读取数据。
    """
    if os.path.exists(file_path):
        # 读取Excel文件
        data = pd.read_excel(file_path)
        # 确保有 '时间' 列，将其转换为时间格式
        data['时间'] = pd.to_datetime(data['时间'])
        data.set_index('时间', inplace=True)
        return data
    else:
        raise FileNotFoundError(f"文件未找到: {file_path}")


# 策略函数
def moving_average_strategy(data):
    """
    12日均线与26日均线策略：金叉买入，跌破26日均线卖出。
    """
    data['MA12'] = data['收盘价'].rolling(window=12).mean()  # 12日均线
    data['MA26'] = data['收盘价'].rolling(window=26).mean()  # 26日均线
    data['Signal'] = 0

    # 生成买入信号：12日均线向上穿越26日均线
    data.loc[data['MA12'] > data['MA26'], 'Signal'] = 1
    # 生成卖出信号：价格跌破26日均线
    data.loc[data['收盘价'] < data['MA26'], 'Signal'] = -1

    return data


# 回测函数
def backtest(data):
    """
    执行策略回测，计算总收益率和最终权益。
    """
    capital = INITIAL_CAPITAL  # 初始资金
    position = 0  # 持仓股数
    equity_curve = []  # 记录每日总权益

    for i in range(len(data)):
        # 跳过前面的空值行
        if pd.isna(data['Signal'].iloc[i]) or i == 0:
            equity_curve.append(capital + position * data['收盘价'].iloc[i])
            continue

        current_price = data['收盘价'].iloc[i]
        signal = data['Signal'].iloc[i]

        # 买入逻辑
        if signal == 1 and position == 0:
            buy_amount = capital * POSITION_SIZE
            position = (buy_amount * (1 - SLIPPAGE)) / current_price
            capital -= buy_amount

        # 卖出逻辑
        elif signal == -1 and position > 0:
            sell_amount = position * current_price * (1 - SLIPPAGE)
            capital += sell_amount
            position = 0

        # 记录总权益
        total_equity = capital + position * current_price
        equity_curve.append(total_equity)

    data['Equity'] = equity_curve
    final_equity = equity_curve[-1]
    return data, final_equity


# 主程序
if __name__ == "__main__":
    try:
        # 加载数据
        df = load_data(file_path)
        print("数据加载成功！")

        # 执行策略
        df = moving_average_strategy(df)

        # 执行回测
        df, final_equity = backtest(df)
        print("回测完成！")

        # 计算收益率
        return_rate = (final_equity - INITIAL_CAPITAL) / INITIAL_CAPITAL
        print(f"最终总权益: {final_equity:.2f}")
        print(f"策略总收益率: {return_rate:.2%}")

    except Exception as e:
        print(f"程序执行出错: {e}")
