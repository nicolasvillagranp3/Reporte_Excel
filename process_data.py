import pandas as pd
from dateutil.parser import parse
import seaborn as sns
import matplotlib.pyplot as plt
import math
import random


def extract():
    pizza_types = pd.read_csv('pizza_types.csv', encoding='latin1')
    orders = pd.read_csv('orders_2016.csv', encoding='latin1', sep=';')
    order_details = pd.read_csv(
        'order_details_2016.csv', encoding='latin1', sep=';')
    return pizza_types, orders, order_details


def transform():
    random.seed(999)
    global fechas
    fechas = ['01/01/2016']

    def encontrar_tamaño(pizza):
        tamaños = {'s': 1, 'm': 1.4, 'l': 1.8}
        suma = tamaños[pizza[-1]]
        if pizza[len(pizza)-2:] == 'xl':
            if pizza[-3] == 'x':
                pizza = pizza[:-4]
                suma = 2.2
            else:
                pizza = pizza[:-3]
                suma = 2.8
        else:
            pizza = pizza[:-2]
        return suma

    def limpiar_hora(x):
        if str(x)[-2:] == 'PM' or str(x)[-2:] == 'AM':
            return x[:-2]
        else:
            return x

    def cambiar(a):
        if a[:-2].isdigit():
            return fechas[-1]

        else:
            fechas.append(a)
            return a

    def limpiar(a):
        if type(a) == float and math.isnan(a):  # Cuidado con los Nans
            return a
        else:
            return a.replace('@', 'a').replace('3', 'e').replace('0', 'o').replace(' ', '_').replace('-', '_')

    def quitar_caracteres(a):
        if type(a) == float and math.isnan(a):
            return a
        if a[-2] == '_':
            return a[:-2]
        else:
            return a

    def add_pizza(a):
        if type(a) == float:
            # Fijamos que sea una pizza random mediana
            return random.choice(mas_pedidos)+'_m'
        else:
            return a

    pizza_types, orders, order_details = extract()
    # Datos que no me interesan para los ingredientes semanales.
    orders['time'] = orders['time'].fillna(
        '12:00:00').apply(limpiar_hora).apply(parse)
    orders = orders.sort_values('order_id')
    orders['date'] = orders['date'].fillna('1000').apply(cambiar).apply(parse)
    orders.index = pd.Series([i for i in range(len(orders))])
    # Supongo que cuando hay -1 y -2 es porque querian poner un 1 o un 2. No tendria sentido tener un pedido con cero pizzas.
    order_details['quantity'] = order_details['quantity'].str.lower().replace(
        {'one': 1, 'two': 2, '-1': 1, '-2': 2}).astype('float').interpolate()
    order_details.loc[0, 'quantity'] = 1
    # He aplicado una interpolacion (lineal por defecto) sobre los datos, teniendo antes que remplazar los strings y convertir todo a coma flotante (Los nans son floats)
    # y no se pueden castear a enteros.
    order_details['pizza_id'] = order_details['pizza_id'].apply(limpiar)
    mas_pedidos = order_details['pizza_id'].apply(
        quitar_caracteres).value_counts()[:10].index  # Cojo las 10 pizzas mas vendidas
    order_details['pizza_id'] = order_details['pizza_id'].apply(add_pizza).apply(
        limpiar)  # Relleno los huecos con las pizzas más pedidas
    order_details['tamaños'] = order_details['pizza_id'].apply(
        encontrar_tamaño)
    orders['sem'] = orders['date'].apply(
        lambda x: x.week % 53)
    return orders, order_details, pizza_types


def encontrar_pizza(pizza):
    if pizza[len(pizza)-2:] == 'xl':
        if pizza[-3] == 'x':
            pizza = pizza[:-4]
        else:
            pizza = pizza[:-3]
    else:
        pizza = pizza[:-2]
    return pizza


def get_revenues_per_week(orders, order_details, pizzas):
    semanas = [0 for i in range(53)]
    contador = 0
    precios = pd.Series(list(pizzas['price']), index=pizzas['pizza_id'])
    for i in range(len(order_details)):
        order_info = order_details.loc[i]
        id = order_info['order_id']
        semana = list(orders[orders.order_id == id]['sem'])[0]
        pizza = order_info['pizza_id']
        cantidad = order_info['quantity']
        semanas[semana] += cantidad*precios[pizza]
        contador += 1
    for i in range(len(semanas)):
        semanas[i] = int(semanas[i])
    return semanas


def usar_reportes():
    # Un grafico de evolucion temporal nunca viene mal.
    orders, order_details, pizza_types = transform()
    pizzas = pd.read_csv('pizzas.csv')
    revenues = 0
    precios = pd.Series(list(pizzas['price']), index=pizzas['pizza_id'])
    pizzas_vendidas = 0
    for i in range(len(order_details)):
        cantidad = order_details.loc[i, 'quantity']
        id = order_details.loc[i, 'pizza_id']
        revenues += cantidad*precios[id]
        pizzas_vendidas += cantidad
    revenues_per_week = get_revenues_per_week(orders, order_details, pizzas)
    revenues_por_pedido = revenues/len(orders)
    comida_cena = orders['time'].apply(lambda x: x.hour > 17).value_counts(
        normalize=True).rename('Probability')  # A que hora del dia se pide mas
    dia_semana = orders['date'].apply(lambda x: x.day_name()).value_counts(
        normalize=True).rename('Probability')
    mes_año = orders['date'].apply(lambda x: x.month_name()).value_counts(
        normalize=True).rename('Probability')
    dia_semana = dia_semana.reindex(
        ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])
    mes_año = mes_año.reindex(['January', 'February', 'March', 'April', 'May',
                              'June', 'July', 'August', 'September', 'October', 'November', 'December'])
    cantidad_pizzas_pedido = int(sum(
        order_details['quantity'])/len(orders['date']))
    pizzas_pedidas = order_details['pizza_id'].apply(
        encontrar_pizza).value_counts().rename('Pizzas pedidas')
    return mes_año, revenues_per_week, revenues, \
        pizzas_vendidas, revenues_por_pedido, \
        comida_cena, dia_semana, \
        cantidad_pizzas_pedido, pizzas_pedidas

    # Dinero ganado -> Grafico de evolucion temporal viendo el dinero ganado por semana
    # Orders hechas
    # Cantidad de pizzas por pedido
    # Dia de la semana en el que mas se pide
    # Pizzas mas pedidas. Esto ya lo tenemos.
    # Pedidos por meses


# Para la parte de ingredientes: Gráfico que muestra como 20 ingredientes son el setenta y pico por ciento. (Es la informacion mas valiosa que puedo dar).
# A parte del informe de calidad de los datos.
# Prediccion. Pero esto como tabla.
# Para la parte de pedidos : Comida o Cena. Dia de la semana, mes del año. Cantidad de pizas por pedido.
# Reporte Ejecutivo: Dinero generado, pizzas vendidas, evolucion temporal, pizzas mas pedidas,orders hechas. Hecho. (Queda ponerlo mas bonito)
