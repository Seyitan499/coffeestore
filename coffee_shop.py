from sqlalchemy import create_engine
import pandas as pd

engine = create_engine('mysql+mysqlconnector://root:@localhost/coffee')


def task1_total_items_ordered_revenue():
    query = """
    SELECT i.item_id, i.item_name, i.item_cat, COUNT(o.order_id) AS number_sold, SUM(i.item_price) AS total_revenue
    FROM Orders o
    JOIN Items i ON o.item_id = i.item_id
    GROUP BY i.item_id, i.item_name, i.item_cat
    ORDER BY total_revenue DESC;
    """
    df = pd.read_sql(query, engine)
    return df


def task2_item_profitability():
    query = """
    SELECT i.item_id, i.item_name, SUM(r.quantity * ing.ing_price) AS total_cost, (i.item_price - SUM(r.quantity * ing.ing_price)) AS profit
    FROM Recipes r
    JOIN Items i ON r.recipe_id = i.sku
    JOIN Ingredients ing ON r.ing_id = ing.ing_id
    GROUP BY i.item_id, i.item_name
    ORDER BY profit DESC;
    """
    df = pd.read_sql(query, engine)
    return df


def task3_sales_per_hour():
    query = """
    SELECT HOUR(o.created_at) AS hour, COUNT(o.order_id) AS total_orders, SUM(i.item_price) AS total_sales, 
    (SUM(i.item_price) - SUM(r.quantity * ing.ing_price)) AS total_profit
    FROM Orders o
    JOIN Items i ON o.item_id = i.item_id
    JOIN Recipes r ON i.sku = r.recipe_id
    JOIN Ingredients ing ON r.ing_id = ing.ing_id
    GROUP BY HOUR(o.created_at)
    ORDER BY hour;
    """
    df = pd.read_sql(query, engine)
    return df


def task4_staff_hours_salaries():
    query = """
    SELECT s.staff_id, CONCAT(s.first_name, ' ', s.last_name), SUM(TIMESTAMPDIFF(HOUR, sh.start_time, sh.end_time)) AS total_hours, 
    SUM(TIMESTAMPDIFF(HOUR, sh.start_time, sh.end_time) * s.sal_per_hour) AS total_salary
    FROM Staff s
    JOIN Rota r ON s.staff_id = r.staff_id
    JOIN Shift sh ON r.shift_id = sh.shift_id
    GROUP BY s.staff_id, CONCAT(s.first_name, ' ', s.last_name);
    """
    df = pd.read_sql(query, engine)
    return df


def task5_dinein_takeout_profit():
    query = """
    SELECT o.in_or_out, i.item_cat, SUM(i.item_price - (r.quantity * ing.ing_price)) AS total_profit
    FROM Orders o
    JOIN Items i ON o.item_id = i.item_id
    JOIN Recipes r ON i.sku = r.recipe_id
    JOIN Ingredients ing ON r.ing_id = ing.ing_id
    GROUP BY o.in_or_out, i.item_cat;
    """
    df = pd.read_sql(query, engine)
    return df


def task6_busiest_shift():
    query = """
    SELECT sh.shift_id, sh.day_of_week, COUNT(o.order_id) AS total_orders
    FROM Orders o
    JOIN Shift sh ON HOUR(o.created_at) BETWEEN HOUR(sh.start_time) AND HOUR(sh.end_time)
    GROUP BY sh.shift_id, sh.day_of_week
    ORDER BY total_orders DESC;
    """
    df = pd.read_sql(query, engine)
    return df


def write_to_excel(results, filename='output.xlsx'):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for task, df in results.items():
            df.to_excel(writer, sheet_name=task, index=False)


def main():
    results = {
        'TotalItemsOrderedRevenue': task1_total_items_ordered_revenue(),
        'ItemProfitability': task2_item_profitability(),
        'SalesPerHour': task3_sales_per_hour(),
        'StaffHoursSalaries': task4_staff_hours_salaries(),
        'DineInTakeoutProfit': task5_dinein_takeout_profit(),
        'BusiestShift': task6_busiest_shift()
    }
    write_to_excel(results)


if __name__ == "__main__":
    main()
