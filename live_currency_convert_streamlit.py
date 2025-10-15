import streamlit as st
import requests

api_key='af929e15595181204f12ef43'
url=f'https://v6.exchangerate-api.com/v6/{api_key}/latest/USD'

def convert(currency, currency_value):
    response=requests.get(url)
    data=response.json()
    conversion_rate=data['conversion_rates']['EUR']
    if currency!='USD':
        result=currency_value/conversion_rate
    else:
        result=conversion_rate*currency_value
    return result

st.title("Curreny Convert:Usd to EUR")
conversion=st.radio("Choose the convertion",("USD to EURO","EUR to USD"))
input_value=st.number_input("Enter the input amount:")
button=st.button("convert")

if conversion=="USD to EURO":
    if button:
        euros= convert(conversion[:3],input_value)
        st.success(f"{input_value} {conversion[:3]} is equal to {euros:.2f} {conversion[-3:]}")
else:
    if button:
        dollars=convert(conversion[:3],input_value)
        st.success(f"{input_value} {conversion[:3]} is equal to {dollars:.2f} {conversion[-3:]}")