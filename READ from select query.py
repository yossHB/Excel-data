import sqlglot
import sqlglot.expressions as exp

query = """
select
    sales.order_id as id,
    p.product_name,
    sum(p.price) as sales_volume
from sales
right join products as p
    on sales.product_id=p.product_id
group by id, p.product_name;

"""

column_names = []
def find_selected_columns(query) -> list[str]:
    try :
        for expression in sqlglot.parse_one(query).find(exp.Select).args["expressions"]:
            if isinstance(expression, exp.Alias):
                column_names.append(expression.text("alias"))
            elif isinstance(expression, exp.Column):
                column_names.append(expression.text("this"))
        return column_names
    except:
        return []




print(find_selected_columns(query))




