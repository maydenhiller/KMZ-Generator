# app.py
import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re
import xml.etree.ElementTree as ET
import math
import base64

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). Produces KMZ matching Earthpoint behavior.")


# -------------------------
# Gibson Integrity logo (embedded so every KMZ includes it automatically)
# -------------------------
GIBSON_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAfQAAACyCAYAAAC0oD1PAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAI9hJREFUeNrsnc9vI8l1x0vcuS/3L9gWjJyHQi5GjGCaAdaIEQdD5egchgySIMjBIx6CbOxgRAlZYC8LSUAW8Y8g5PgQn2JpYHvtwEDUcwkMxLE4cGAEDgxx/oLh3kfD1JNez7Q4TbJ/vOqq6v5+AGI0EkV1V1e9b71Xr14pBQAAAAAAAAAAAAAAAF7wf7/9bYBWAC7TQhMAAEAmOlrUO2gGAAAAwH8vfaRfbbQEgIcOAAB+c6xfR2gGAAAAwH8vvfO/v/nNLVH/m3/8YQ8tA+ChAwCAR/zOF74wvbq6+vx/fv3rfvy9rS0V/e2nPzzSL4TjgTW20AQAAJCfZ7/61cWrV68GO3fvTun/f/dPP2ovlNpTC3X28V9/dYoWAgAAADzgvy8uwl/88pcX//WLX9zyyr/5rR+Pv/GtH/fRQqBqEHIHAIAC/O7OTnT18iWF32+tp3/0V380eGdrq/333/5sjFYCVYKQOwAAFOQ/f/7z9qtXry4Xi8Xw97/0pUnyZ/vf/Ul/sVAP9Zfdw7/8yhytBeChAwCAo/zeF7841176iX4d/UcU3So6c/AXX5ncaW2dvNPauhh99ycoSAMg6AAA4DJXr14dv7y6ml9dXb0VYn/053+oRb118E6rdX74zz/F1jYAQQcAAFf5gzAkL/1Avzo/+elP3yo6880/+/LkzjutoX6d/sO//PseWgyYAmvoAAAgwI8+++xSLRbBYrHo/vFXvxot//zjyc/6C6XG+ueTbwy+PECLAXjoAADgINpDf/zy5UulX+N/+8EP3iow82H/g4n2oIatra2+FvdT/UIRGgBBBwAA5wT96uqYBF3/G+jXftp7tKhTLfiJftF6+jlEHUDQAQDAMf5kd5cS4yYs6nv/+v3vhytEfcCi3mFRRwY8gKADAIBLvOSw+9XNa/z4e99b5YEP9WsKUQcQdAAAcJA//drXqHrc7OXVlXq5PvROhWa6+kX/tiHqAIIOAACueelXV2cJL33v29/5TmeDqKuEqGOvOoCgAwCAE4L+8uVTznYncVfLtd6XRJ3C7sOEqFP2ex+tCIqAfegAACDMp59+utCoBf2H/l0sBl//+tcnq95P29jUTeZ7zIC2uaElATx0AACw66VH1945e+n6dfTJJ5+s26JGme+zxP/H8NQBBB0AACxzdXU1vYrD7jdr6W3ayrbq/byevlw9DqIOIOgAAGDZQ3/+eg39jbA//Oijj4I1oh7pf44h6gCCDgAA7gj6NM50f+2lX121V21jS3CgbofeIeoAgg4AABYFnQQ8Kebx1/1Hjx6t89LTQu8QdQBBBwAAG7x69WoaC7n++vpFWe/MWi+dQ+9nKT86QvEZAEEHAIAKGY1G8zU/XuulM8OU76GiHICgAwCAY2zy0mfq7QQ5iDqAoAMAQJVoDzzc8Jaefs+mo1MpQW6+QtTHOHoVQNABAMA+JMb9DV46ifnJih/Hp7RB1AEEHQAADBJkeM/DDO85XuGlx6I+RlMDCDoAANgV9GBTaJ699LM1b+lpLx2iDiDoAABgiHsZ3/cgw3sONvy8jz3qAIIOAABmCDO+b+P555zxHm142xiZ7wCCDgAAgjx69KiX4+3tjO9/nOE9lCQX4AlA0AEAAMhwX/r9fDb6fNPkQL9O0fwQdAAAAOW9cxLVXs5fCzO+7yzDezpIkoOgAwAAKE+PPeU8BBlKwRJPMn4ekuQg6AAAAEqyX/D3NnrpH/Y/OMvxeTjIBYIOAACgCNrLJq84KPjrdzO+L6uoozwsBB0AAEDF3jmR1Zt+lvMz9/FYIOgAAACye+d7JbzzPIIe5fzcPe2l9/CEIOgAAAA2i3lbwBPOFBr/sP9BVOCzEXqHoAMAAMgimCp/ZvuqiUEWpgUmC9jKBkEHAACwRoQpnC0V0s4adp8V+OweQu8QdAAAAOliHljyfJ8V/D2E3mvOHTQBAGCZxWJB3mJs/MOUt0RL/59vbW1NG9ZMp0og1F6AWcHfo2s90q8BejgE3ZZhoVlwkDAqyWMJwxwfRcZmnvj6cx4Y1y9tjGboDqCh4h2PLxpbHZUt9Luf8jmvxV29Wed9mjL+VuHNpEB750cqe4jcFUEnqIrc44LJdUX61jm3U7eBE75mC7p++O0lwxIKfnxn3USAjVHEg+U5fz3VnXDuWBv1VbntMUYMjG6nSYZrD4WfqSS3DJy+n6jOA5/HGq2pPjQgTO3Ecw5zPoOuB2JOY3DP48dPywTbFf2tp9wHznWfG+hxdVaDsRM4aIOvbdYdBxqHjMkDfui2SxWGyx6Ivr4ZGxrqmJEDnvwDB0XxOEf7ulroYn+Ftxl7lk/566nP0RwW8j0WctfWU5334LSYd5ShdfPDw8Osk8iy/S/QXvpIe+mjCppsnpjknbKoT5Tf9B21Y1tbFmc4D9lDCDx7mLHAP2GBn1tsQ3rd5w7WrvD+T9j4Fl6qYG89cHSCkrkP+ORx6DbfY0PkamLUgW7Pkavtx0lwF6baTwt6ZnusBXkhILTbWtTnhvvctXe+9G2vRT3FQw/ZDlflkJLtecz299Yk8E7FDdH31IDfmt2ygPb5ns5Y3M+qFHcW0mth0ddwoG6SXfq+GNxER5zo6+8pob28VfcBfe30zKkPnLi6RsgGaOz5uLMt5vFZ46b6aNV9p6oEubT7GlMEzFdRT9jepMCOKrBjc54MrXQitioyKHGIwpQ3HvHNrtvOESfTdQw1uHXDrtt5bFDUjc6qeenl3CNRX9UPD1xafxcyMlO+t+cJA32dX7LkrdAzfF8VXz7rupi7wGJ+btgDm2gPPbO4CnjoMeSlzwz3wYUNm2JpvMVLMtJ9Za4yJBbeMXxzJoR8pkquafM6Ypx0d0/Ic2knvDa6vmHVwq7/3iCRsSztmU8MX/uU1tfYC/IVaveQn//A9lo7j79xCQNCSyuTdfex5K1ES1GBOOkuUH5TRUb7M0v3Rv3DdCLidEX7kaceuLzMUsKOSTsnmfRky5AhCYUHgXHvlz2Z+4IerpX1QDakl5ITKH0f2xVe/7nghIT6yjDl+3GU5i7/LVMRG+oDxzYMS0kxpyWcY6klpBwT+x3Xli20d24y6nXr3rWHnune+ZzzC8G/3TW5jS3DmKZJY632xut7Him5xDlyXDNNulrCN9HWryMlF56as3HZpgducrDTugR3qvf4b8597EjsMUkmalWd9HUiKaoUwk15kViN9GtXv+h575BREb6P6zVKXgap2piEBcV8zqI6kswH4ejOxjZusJjPsop5om9JYjtju29jnBhGciL/OOsbW8JGhGaNUvszJyzko4qTzebsWW97LOyS4bsnFV97VHVjkZDwZG7bwASmUmPFy0lFli0yrdGVHFeDFRET59BiPqpIzG1MmpcJtdcfGvz8qWvjpAodEbRlmT+nJWRERuyVBxKzVTYsA5tFXZaEfeJZf5p7PhCsRTfIa1fymb9VGquiCXCVVPLiJYjjFePeFTHvV+y1Ps75fhPr+Sbv9/MGe+oidqkSQecQ+7lgZ6CZ6o5Lma4Jz2LXI6FEicVyz3xiSNSNVhfjKFmRE7UqTeDUf2uY4nU4Iegs5lWKSt5wu1Jmcj5CXpu3DY2TU440gZy0ShiPeJtRKHQtE17TdFI0ee/fDsSyUaJ+IO0FcdKiKYpMrKeWEvecS4KyIOaqYB+7a+haHjryKGhSeg5Rr0jQE2IuNaMb+JDlyKGPLkS9MaI+En7WbWUotMljssjk+sTiWJokJxaWxbxnQczj3Tt5MTUppINbAkeGXweiXoGgGygA4lVxAY4gQNSbg3QSV9+Ql/6wYH+2OfaSk4nPbV2Eyfrsm+7/8PCwSETSZGi879DYg6ibFHQDYj7xsVIQRL1RXnqk5Nd3ewYutchnRpbbdqosr52zmNuoUEg2JPdSh+Fs9MITQ4i6Z4Ke2A4j1bBnPhcTYFEfKI8zykEhT1KCB5IfxhPtIuNy5kDbTm1NLviwFVvlhg8KeuemBb2tJw09x8ZfLOoBTJGchy61LS02JN5XBmIPY4BuVHukxaYj7HEUNfLPHWhbKyVPKzhsZa3902JeNBHxbgXX98DBMXhdHY8nr6CMoPPeQMmGdDabvYCoU1LLGbpSfeGJm3R/lRxPRY38PYeauWp7YPqwlXWUcQKq8J572kt3McTdZk8dol5U0LnGeV/wbx64esxkyQGK0Hu9ke6zoeBnBRX/nuRkiSpBblVpE7ikqy1RONPeeVTkFysOhfcdHYcQ9aKCnjhHWcwo1ulknYRRik+mAhB0GxQWdC5G0xgqLumaFoUo453fr/BaHzj8GCHqBT106cPahzVuy2N46bVGeluVZLi7jKe935QHyHvNbd5v0US4mConXx2H9qSvE/VGTUgLCzqH2iUba+JSSVd46QDciITpkrSOiHmg7Ow1j4lKJMLF4faqBdZ1sYxFvY9hvNlDP5KenTagPY/RpWrLrMb3dtSA8KWtjHaibKiduG/huu978mzHEPU1gs6NIzkbnOQ5McZzL32CbgVB93CCUNs1Se2dk3Ni894o1F74eXPGuQ3B6nn0mCHqazx06XWmJoWinyhQR6QF4blj91fL8CWvm9tcUigVamesPZMKKtNB1E0KugHvfFrDbWrrvHTak47kuPohHa6VHBMzwXsko3hUh1KbXDzG5rq5RKidsFmONfTssY+bkBOSx0OX3q7QxESxSFVT0Qn4i4uCHkMG8aIGGcRHyt66+bXtKxNqZw/ZRjJckns+PncuhtZsQed959KDuHFV1Phc911oVq2QnKDNhHNKTETAyBZQCP7UR29de+dkx/o2J2xazEeee+eEr3kV/SaKestw54nqUuIVNB5JL0l6kvvU4H2Th3jJ21h9885tUrrmBq9fh5bvgw5rgah7KujSgxYJYsB72EOVNGrSy1CRaaOuX6e+eOvaO+9b9iwnRcu7LuFK0Z/A4+HbKFFvJYxWx8CDixQA/iM50RXfwslRsLOK2uHSg6Qjm0I4F/LOpQt7lcEVD512CxRZXmqMqCc9dOnOM29SdjuoNZIFNkwVWKoqGkYeOiUdOXk+NXvnNq/rpGR515gjh5rVlQRfKr/cLegokqhf1GH3RlZBl85mhJgD72HRkvLQj00VWNKfO1HVFqshB+DCQW/dtndeulqk9s73lFthbmdEkKJR+kWiPikYaTivs6ib9NCfKgD8R8pTmirz5Y+rLq+c9Nath2W5iIzX3jlXhXPtwJzQtUGpRX0AUV8h6OyFuFw4AwAb3nlfyDu/LjBieseHBS992VsfWX5kNo/8FPHOlfwJl7UFor7aQ+8Y6uAA+CrmHSHvnMZBt8J8koHFZtvndcrKvXU+Tc3m1rqJgHceKkfrp3PkAKLeVEGv81GpoBGe+bmApzStWMzjcWfzZMOOJW/dthCW2orIgulyJraze9FZ1I8L3tNlnQ4ligX9fZhxACFftKmOuZIJex5XLeYJAzdS9k/9I2/9ssLysTaP+pyWLfGqbtbNA4zCwn2etgoWiU7FhxLVQtRjQZfuSFg/Bz565Req3Mlc8fG522RgbFZJLBGKlCRgY2n0sBc+hCW0eJ+PS3rnobJ7IlxdRH3SdFFvGfpcrJ8DH7zxHovNJXvlRSe2VNSF6ve/R0Jqamuap6KulPnDXkLL91e2qM8YIxKiLinoIboCaIIXzlusFvq/L/TrlMUmKPnRtH57yp89cskosKgPHbiU2FsfG/DWbbb3rEy4XXvnI4VQO0TdcQ8dABchw0v1EQ7Yq4qUbDSJJsb77I1esrhbz6LVBo7W83eUG0thfSV/2IvNIz7LtukDDEujop53fLd5/PYh6HKdHAATg5xO/xvxi0LklLT2nv7RNg/+iaDAByzuL0yvIWe896l+7fBkxvaSWPKwl0CorW3xrIR3brsQThNEvVuwv499FHVTgv45uhPwaODT+eQTDk/H4j4T/BN7ypEjSDkDfke5cXBST8mUj7UpimXa0SfvfOrp2J42SdQRcgfgtgGYs7hvK9n93LFXOnbAW59xPeyhI956XD42d7s8evTI5ySm0JcL/bD/wdzjMd0YUb8DE/669K0rDy1CUR5nDAGtgdNaO2UhSwkH9bOO/tyuzW1tfH/HifuzLS7090nUBzn37tvOUSjkuX48+VlH+VPidV6DsTylMaeKFYwiUY9D+BB0DyBBd+lABAi6m4ZAStTjamq7to8Y5i12XQ5771sWmbgcZ9eXo5dLlHv1KbIwreFYLiLqd7mAjbPEIfdZw402ecWvUTfrqEPD7UKGgEK6O1u3GUFGnesf1/XYhQ0bTSJPXaklnciEtz2ZbLPxrPsBJQEE3Y6os30vck97tGTWREF/3/OHPksYuImhAbLN2dbYEeCPqO8q2fBjoBw6ICKxtj5QdsOs5L2e1rxLvevRta7N5Pctj6HkBL3vsqibSooL6jDi+MFLe+rx6VuopufhRE/Jn2ZGxnDfsfucsBdzZvEyQoHsd5epU8g98HAs11LUY0F/CnO99sGfCH7kBGLudX+IC9JIslfhISaZ+z3t1TcQlcjDvuuhd+2dhjXv8rMP+x9sEr25p2O5dqLeMvRA6rb+JRkWxx59/zFxPOmRizfKExhb3nrb1XZxzKaYJMvkNfD1IQiIulNnqrcMda6OAqC+XnpkwEvvuLrf1bK33ttgMG17h0VtnS8T+yd1d+C4fxfNlwqVQ3kwpgRdNSBLFTSbxwY+0+nKYZa8dbIjKyvsHR4e2vZ0iyYAzzzo4xRuz/Ks67KtrejphB1XRL2VCDtIdzB46aDORAY+MxSqbV43b/3+hp/bzsgvgg8imHXSGtZlUPsu6i2DBgqCDmoLZ7yb8LJ6ntx/ld76pjaxKY6FxCxDopltaJJ03NCx7a2oJwX9ifBn34XZBzXHhFG+78vNL3nrM5N/a8MuAKviWGIf9pnDj/fE5/rtlkX9wtaZ6iY99FABUG+eGfhM7yJb7K3HR7OaIqj4OVRh6544+kjzeuf36ji4WdQHBfvqubKQ/d9KzraFZ4yB6+uBAAgYPmnaPo4b9tZHLOwmPObAVQ9dFY+qnCk393AfZPXOHz16VGsbz0WWioh626qgG5oxwksHdcaUkHhrJKmUcYktQIXgTHebwhhqYcu9bsqi6VrYfaqvK493HqoanMZmSNQrp5Vy4ZIP574CABiH1pi5yIUTIfsS4cpVbKp9blsYiyYzHjjWlfI+Mwq3P6v7+PJF1NNquUvOrHvYjw5Afo+vwO/s8+85M96EjWCw4ee2y1c/LPJL2hueKXeyyQ8KZN/36u6h+yTqaYJ+Ivw3+rDPAJj1zpWjy1tsBCUEa5MXaNtD75TIdj9wQBQjLeajPL+g77fHE8jGnBjJ/bnr6iSmlXLBM2Ev/aECADjnHVZoBIemjf7h4aEL69FFvfS5Zc8vPhq46P026ghoLv3spKi31swYpQhcO0kKACFM9evMBpIz4n0oRjOs4G88tnyP/aJZ31xi1Ubo/fpwkrx7zvk+qf/PeDLVKCj500VRb6242JmwqO8rAEAeI5uVZTEPXLwhgQNtZhm8dBe2gRW2dVpUhxaiDLsFq9bFp+BFTR2kLop6a83PjpVc9acQXjqoIaYKauQxEMs7SVyu0FjGg85qi04s32O/xFo6QaH3KkLYcxbz3ILMZ8DHE8mnTTYArol6a82F0gVKhsngpYO6YcQbZiORebK84f8uMa3gd48dMK7jEl56fD73xLCYdzOepLYs5u2l+4uabgR4vJoqqCTmocclHaU6Vujqec8A5IXXrk0IepTjGtLEu+PqVtGcE5Vb3jk7GBtxJDmOMt5HZURdv8hTN7FH/Vp8ShwOc5To91Pd3jNYg9fL1F3bot7K8B7JDNUj7EsHNcFUIlqeEGZQ8bXZIq8X6MI2sP2SoXfF28ikPD9qD9pnvsN733Oj74ccsqRT9hhm4Jaoz22LeivjRQ6EBgiJ+SkePagBDwx9bh7vcpWg122raK6S1Ow1njhw3edFSsIuiTqVYt1hG1xEiMluT/RrO+8+8yUxp8nJ8lLCBGbALVFvZbzIqZLbJ0mh9z3PnlOArgpiuLyqiRKrUc6w9KoEuI6LSagFD52Z89JfXiSTess4MKVFnYV9ol/bLBbHGwQjXnYYsJAPyhyFymJ+vizmTdyu5rqo38lxkWd6QA5UiYSPBBR6n/JWliZ7Y8BPTCV45g1hrhMKWuvccazdikyCCq2Hk9hoIaLlwlMH7plEvSshgJyV/tpufjz5WWepH0wlzzFPiHm7ZF9tnKhrjevyOOxX9ncLzLL7QqJ+PYspkShTlVcRpsxOy3DAx0z6cs3dqide+voXgh5v1/H+8NoQ8yllea6FrmOdJz7Un3ns0FgaFzBu25xwVFSQSNBdyCm43t7kk1fLa+ZHKWIe6fvo5njuI6FJsHHb6VC/T04OMut0q8CHT5RM+P06HOXK6VArHkRbaPICaoDh/mCiktq+K+OL2y6vsE7KiDkjlf8j5akHnoj5Eff1tCjQAaxBLs2kPjip4m+1Cl5gU0SdZpXSA/B9dHFvOVVm8ikODEVBricgjuws6at8J8HNJYSDPWJXTsgiO3fBh5q4KuS05e5Cf7kqz+lMt2kEU+CmqLdKXKC0qIeOeWP9NZ26DAG6t3+eOYfNTPTRyHAYMS2hyYZ3njfkeiDgnceibqtO+ip7d0oesESynKCQt3nv/IVan+swhEVwV9RbJS+QLm5HlQ9pxaLuRPa7YJ6ALUHHXn+5vhCwIPYNfDytq+5WcBuU9X5q0VM/zdknz6TX/rWoGz/xLSdk6y55ndq2mPdZyDdNug5QSEZE1HNFnvKM2zsCFzjVf3CbB21ZD4ay36k29UBqdl7AgB8Z8sxfCzo9oKyVr0p4ZaB8X9hjI2dCCK+TpEr2g2mOMdfjvtc13PeW2zBvZENyi+wyXRauwJEudr0kogV1n8VyUtUf5ugACfnDjO1BVeFGsAoioj7S42KWw2kkex4Z99ATFzjnbGKJZAka/BecGVml8ab98ReGxbwqwX3fo2tdfg5WJyMcXt/Tr0uVnuErASV77QgI67MCz/JS31uvona8yBnZkJjkrPPS43O/Xcs0D1jYLzkUb2zCQev3+kVC8kLdLuO6jrJ5CO8KXf67qiYILlnf/lxDBnksJAQzniScmRrkHFLdVxXuFVQGt19weOZSUIjEt35tuP6RktvnvfHaub+2eSJ5T5k93OTaMBYslLKq714WbRtlKBLGE4Zxzj5oVMyXRG3V3mqXmPIzesLe8bzEvSb7dpF7HpSJHvDkWGKSQjX9t+vkrWdc3s28dXjL4IVKhivjykcnUvvW2ehQwRgbGafGRFJYEHN3KMcmIy5xzBO5uQWDsEnYTyQmGTymHhaYFJFYDKtcCvBE1JedG3o9TdjE6ZKXH4vm3cQktXS/5fwDW/3Tii1yTNQH7NHbE/SEgd7jQS41cOZshJ5yh55uMgR8HXG5zjIzVUneM2Dcr7fFGPIsjRcBosQtVb+DRSZKMGN7zXMvu6Xu1rjKYjQTJXDv8XNrF/ibYhGLBoh65X1Xi/mgRL80MUGncbRT5eSvIlFfF9XKHNHdquhiTQj7qhlskljIXWRX0pCxcTVpnIyKetlqSo4RH4hxUlVyZ2JrmHQOyFS9vebcEehnlXvlEPXcz71wZTvuj+eG7G9lyzMVi/qqvjjh7Hg3BH3pIcdhOd8ysckAUf3i+0JGk/ICdoXaNF7eqIJjSaHi+gNHyv/M/HhZ6Iktj5PbM1DV54Q4FbEoIOoBRziwO0RGzPvKXEJpcrwNs4aiPRf1zEu0W5YvPF7DDhxt32sDrRJJecJh7dxhdx4s1F7vsgEKLbVNxAP/c+5wUUbxjq/3nrq97ucbM77/pyr/KWlVTZ77PMZsC1XlEYsCoh6X9a3bkk+Rcb2bRcx58hhPHGN7JBG9ydu34nGY9Ghnvj6AFFGnXWTvOS3oKTcQKvvr23Fm6dN1XhZvx5EwkrkPz8hwIIeViU+WaIOhhD2TRiIp3s/Vm2WdqU/hPja8vYrH14zHktWIRQFh96WPGome5FkzN3hQUVm8T5xLEfVMzt+WwzdDr7uGZn2xwSaj80xlTAJa8pIlMjdrtw0DeCPwyTEWlJygxuNpyuMp8tlD4nV1U3X7XWWoxfwYo8M5HYy3gGeapGx5eIOxsOfxUqNE6EJq25vU3spB3daBgNdGZHlcBYl+HkcokhPSWR3bgUPwJpIMXeO62A4OXHF2PMY5UlHtBN2xhh4pmdAcvHQA3BV2muDUIWkzDVoKGfh0RjuAoJsUdSkvfSh9GAUAQFTY+8p85nZVzNRNiP0MTxaCDt4IOg1yibV0miFv121fJQA1E/Uq6mmYhOzLibqp/gZbA0EHKaIulXUusi8dAFCZsNOWwABCDiDo9RF0GtAXQjN2JMgB4Je491nYQwcvb6ZuimFByCHoIIeo02z9SGgm3XWtSAkAYKOw08Q+PvDJZgLd64qFWCOHoIPioi51sEgt6xQD0EBxr6qQD3niEUQcQNDlBF3yMAKIOgD1EXjJQlkz9eYY1etiPlrEZ2hlAEGXF3XJk5sg6gDU35MPEt8Kl94SJYUcwg2ABVHXrxcLGS7Y8wcAAABADUQdxzoCAAAANRD1F5xJDwAAAABLon6xkOOc970DAAAAoGJRb7MQSzLC2joAAABgR9hHwqJOYfgjeOwAAABA9aIe0gltC3koAtCH1w4AAM0F+9AteevK3IlNtH/9ibrZwzrNs4+ds+nJ46d/I/27EZ4WaMB4pD7fRsllAEBhI6Jf40U1XLAXv+qVxqnU5GX5g21NohKXcLlpK2DKdY8sXm9ewoqvNRcOjL2Qx97liqWsUz4aWTnQV1eNZ7rGPReicmk5PhaeZ+F8pIqusZ/nb6/IvXqBJVY/hP1IcIub1Np8x5RxcsRIri3aA0Gvn6CzkTx1vQZEzmf/wubkA4Ke6zrTdjwFOfrAxmttQVLtsrW1NdOvof5yW78G6uakpKqZ6ddEv3b1tbxH19OA8CMZ6jF6YHMmzurmmONezj5y7vit0aR0XPUkDhRimPK9/RXvfZBio483/YE7aGNnhH3Oojphz5EMz31l7rSmSN0c8HDW4LXDHs169f2PHLy2mbpdy5sI1O3a3yrlPYTt+v9x33JFzGn8nKa0XdzOU26zQL1dT/3Ekds44H/fZduwfC90fPNOQ8fxPGUctNXbB2VNU8bGrEIbH+m+eLY0qaRQ/AE5don+2k95vkOc61Ef76LDazAjXlfJE56/5N854jW3TsXX7mrIfW2I2nbI3eW2TLmuhWttleHZv1jx3NuJ95+7+pzTcl/w/G9dU2h7KWrFdQUp9nu89J7l3I7M/RAeuh/e+5Rnl28J/RrvfYoZXWZoXXUnOUsGteJhyve6aZEpHjMjfrnMyXI0gcQCfdh5Wz7Tz4meXTLUTpHCaw98hXc+gKA3R+hBedos6jiutmawV7Y86T2u49iBmHsDrYU/SAg39c89nkTup/TVzM8Vgg6aChn05PIDfX2UZzYMVnJvRdh1YkF0wpTvPa5h1AFi7s/EizxxSpBLbg1+qL83X/LO6f8HeT4bgg6aymMeMEmDT3kKz/SAO0bzlBbRNCGNXBCeNO983fKVCwWWEuu/9O999XbC1wG6nVeifqafaZQYJ+0U7zx3IhwEHTQVGkC76mYrU3JWTMmDWMpoHkcrJiHX9teB61uXGEWRjwkeoXcM2f4kbVLMtMgzxT500ORZ8pxFfRkKhb2LFqonNTrzYM5eHJaJ/LQ/5DgcrxH73EDQAQbV2+vmZPD7aJ3C0L7aNCJLordMmPF9LkFtN0vpp++ju/k9VlK+Nyk6ViDoAKJ+E9qapBhL4D9plRfTtrHRpK5b1DOqoI/Ste2kiPoeqsR5bXvSJpLPi34eBB2Am4FFBh1r5/V7rjP1dhWxcLn+ORlW9oqmDt/LXKXvwjjCkwYQdABuQ14Q9qHXj7Sw5pgrsi1HYpyOzPCkY3mCQpUk9/CYAQQdgNseUBctUZp9l05bYxFME3XaJvQicYzwpbq9N9hVBivaPEDXg6ADAN4Yfwq5DtEStXuuI7U6ozjkV5ogzh28l1nKvVBkAaF3CDoAYMlgkrGcoCVq91xpokYRmFnGX6GEOldPMDtImWz0kCDXbFBYBpgmcvQ6og3Gf0BV49SbNdUIbblWXHwRdWrDba4MR8dY3lt6C0Vo6LlHFmujRxnug8qHdtXb57sHDjx/F/rpLOW6Zp6Mn0gBAAAAoLn8vwADAIDWnwBtGEB/AAAAAElFTkSuQmCC"

def add_gibson_logo_overlay(kml):
    """Add the Gibson logo as a fixed ScreenOverlay (bottom-left corner)."""
    try:
        from simplekml import OverlayXY, ScreenXY, RotationXY, Size, Units
        so = kml.newscreenoverlay(name="Gibson Logo")
        so.icon.href = "logo.png"
        so.overlayxy = OverlayXY(x=0, y=0, xunits=Units.fraction, yunits=Units.fraction)
        so.screenxy = ScreenXY(x=25, y=95, xunits=Units.pixels, yunits=Units.pixels)
        so.rotationxy = RotationXY(x=0.5, y=0.5, xunits=Units.fraction, yunits=Units.fraction)
        so.size = Size(x=300, y=0, xunits=Units.pixels, yunits=Units.pixels)
    except Exception:
        pass

# -------------------------
# KML namespace
# -------------------------
KML_NS = "http://www.opengis.net/kml/2.2"
ET.register_namespace("", KML_NS)
Q = lambda tag: "{%s}%s" % (KML_NS, tag)

# -------------------------
# Color map (KML uses aabbggrr)
# -------------------------
KML_COLOR_MAP = {
    "red": "ff0000ff",
    "blue": "ffff0000",
    "yellow": "ff00ffff",
    "purple": "ff800080",
    "green": "ff00ff00",
    "orange": "ff008cff",
    "white": "ffffffff",
    "black": "ff000000"
}

MAP_NOTE_ICON = "http://www.earthpoint.us/Dots/GoogleEarth/pal3/icon62.png"
MAP_NOTE_FALLBACK = "https://maps.google.com/mapfiles/kml/pal3/icon54.png"
RED_X_ICON = "http://maps.google.com/mapfiles/kml/pal3/icon56.png"

# -------------------------
# AGM icon whitelist (YOUR EXACT REQUIRED OPTIONS)
# -------------------------
AGM_ALLOWED_ICON_URLS = {
    "http://maps.google.com/mapfiles/kml/paddle/ylw-circle.png",   # Yellow Preliminary Location
    "http://maps.google.com/mapfiles/kml/shapes/triangle.png",     # Purple Valve
    "http://maps.google.com/mapfiles/kml/paddle/blu-circle.png",   # Blue Survey Point
    "http://maps.google.com/mapfiles/kml/shapes/flag.png",         # Red AGM
}

# -------------------------
# Helpers
# -------------------------
def safe_str(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

def normalize_agm_name(raw_name):
    s = safe_str(raw_name)
    if s is None:
        return ""
    if re.fullmatch(r"0+\d+", s):
        return s
    if re.fullmatch(r"\d+", s):
        if len(s) >= 4:
            return s
        if len(s) < 3:
            return s.zfill(3)
        return s
    try:
        f = float(s)
        if f.is_integer():
            i = int(f)
            s_digits = str(i)
            if len(s_digits) < 3:
                return s_digits.zfill(3)
            return s_digits
    except:
        pass
    return s

def normalize_color_value(val):
    c = safe_str(val)
    if not c:
        return None
    cl = c.lower()
    if cl in KML_COLOR_MAP:
        return KML_COLOR_MAP[cl]
    if len(c) == 8 and all(ch in "0123456789abcdefABCDEF" for ch in c):
        return c.lower()
    return None

def set_icon(point, href_value):
    href = safe_str(href_value)
    if not href:
        return
    try:
        point.style.iconstyle.icon.href = href
    except:
        pass

def set_icon_color(point, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        point.style.iconstyle.color = col
    except:
        pass

def set_linestring_style(ls, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        ls.style.linestyle.color = col
        ls.style.linestyle.width = 3
    except:
        pass

def choose_note_icon_href(icon_value):
    v = safe_str(icon_value)
    if v is None:
        return None
    vl = v.lower()
    if vl == "map note":
        return MAP_NOTE_ICON
    if vl == "red x":
        return RED_X_ICON
    return v

def haversine_m(lat1, lon1, lat2, lon2):
    R = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(a))

# -------------------------
# Line builder (prevents "looping back" by splitting on big jumps)
# -------------------------
def add_lines_with_autosplit(folder, df, color_col="LineStringColor", split_jump_m=5000.0):
    if df is None or df.empty:
        return False

    chosen_color = None
    if color_col in df.columns:
        non_null = df[color_col].dropna().astype(str).str.strip()
        if len(non_null) > 0:
            chosen_color = non_null.iloc[0]

    created_any = False
    seg = []
    prev = None

    def flush(segment):
        nonlocal created_any
        if len(segment) < 2:
            return
        ls = folder.newlinestring()
        ls.coords = segment
        if chosen_color:
            set_linestring_style(ls, chosen_color)
        created_any = True

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            flush(seg)
            seg = []
            prev = None
            continue

        try:
            pt = (float(lon), float(lat))
        except:
            continue

        if prev is not None:
            if pt == prev:
                continue
            jump = haversine_m(prev[1], prev[0], pt[1], pt[0])
            if jump > split_jump_m:
                flush(seg)
                seg = []

        seg.append(pt)
        prev = pt

    flush(seg)
    return created_any

# -------------------------
# Placemark creators
# -------------------------
def add_agm_point(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return False

    p = folder.newpoint()
    name_val = normalize_agm_name(row.get("Name"))
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    # AGM icon: MUST be one of your exact URLs
    icon_raw = safe_str(row.get("Icon"))
    if icon_raw:
        icon_norm = icon_raw.strip()
        if icon_norm in AGM_ALLOWED_ICON_URLS:
            set_icon(p, icon_norm)
        else:
            # If it's not one of the four, do nothing (prevents unexpected Earthpoint substitutions)
            pass

    # Tint by IconColor (Yellow/Purple/Blue/Red)
    set_icon_color(p, row.get("IconColor"))
    return True

def add_access_point(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return False

    p = folder.newpoint()
    name_val = safe_str(row.get("Name")) or ""
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    if "icon" in row.index:
        set_icon(p, row.get("icon"))
    elif "Icon" in row.index:
        set_icon(p, row.get("Icon"))
    return True

def add_note_point(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return ""
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return ""

    p = folder.newpoint()
    name_val = safe_str(row.get("Name")) or ""
    name_str = str(name_val)
    p.name = name_str
    p.description = name_str
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    href = choose_note_icon_href(row.get("Icon"))
    if href:
        try:
            p.style.iconstyle.icon.href = href
        except:
            try:
                p.style.iconstyle.icon.href = MAP_NOTE_FALLBACK
            except:
                pass
    return href or ""

# -------------------------
# KML post-process: StyleMaps for Notes ONLY (hide until hover when flagged)
# -------------------------
def inject_hover_stylemaps_for_notes_with_flags(kml_bytes, notes_flags_by_name, notes_folder_name="Notes"):
    root = ET.fromstring(kml_bytes)
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        nm = folder.find(Q("name"))
        if nm is not None and nm.text and nm.text.strip().lower() == notes_folder_name.lower():
            notes_folder = folder
            break
    if notes_folder is None:
        for folder in doc.findall(Q("Folder")):
            nm = folder.find(Q("name"))
            if nm is not None and nm.text and nm.text.strip().lower() == "notes":
                notes_folder = folder
                break
    if notes_folder is None:
        return kml_bytes

    style_by_id = {}
    for st in doc.findall(Q("Style")):
        sid = st.get("id")
        if sid:
            style_by_id[sid] = st

    def href_from_style(style_el):
        if style_el is None:
            return None
        href_el = style_el.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()
        return None

    def href_from_pm(pm):
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()
        su = pm.find(Q("styleUrl"))
        if su is not None and su.text and su.text.strip().startswith("#"):
            sid = su.text.strip()[1:]
            return href_from_style(style_by_id.get(sid))
        return None

    def get_name(pm):
        n = pm.find(Q("name"))
        return (n.text or "").strip() if n is not None else ""

    pairs = []
    pm_info = []
    for pm in notes_folder.findall(Q("Placemark")):
        name = get_name(pm)
        hide_flag = bool(notes_flags_by_name.get(name, True))
        href = href_from_pm(pm) or MAP_NOTE_FALLBACK
        pm_info.append((pm, href, hide_flag))
        key = (href, hide_flag)
        if key not in pairs:
            pairs.append(key)

    if not pairs:
        return kml_bytes

    first_folder = doc.find(Q("Folder"))

    def insert_before_first_folder(el):
        if first_folder is None:
            doc.append(el)
        else:
            idx = list(doc).index(first_folder)
            doc.insert(idx, el)

    key_to_smid = {}
    for i, (href, hide_flag) in enumerate(pairs, start=1):
        sm_id = f"sm_notes_{i}"

        st_n = ET.Element(Q("Style"), {"id": f"{sm_id}_normal"})
        is_n = ET.SubElement(st_n, Q("IconStyle"))
        ic_n = ET.SubElement(is_n, Q("Icon"))
        ET.SubElement(ic_n, Q("href")).text = href
        ls_n = ET.SubElement(st_n, Q("LabelStyle"))
        if hide_flag:
            ET.SubElement(ls_n, Q("scale")).text = "0.01"
            ET.SubElement(ls_n, Q("color")).text = "00ffffff"
        else:
            ET.SubElement(ls_n, Q("scale")).text = "1"
            ET.SubElement(ls_n, Q("color")).text = "ffffffff"

        st_h = ET.Element(Q("Style"), {"id": f"{sm_id}_highlight"})
        is_h = ET.SubElement(st_h, Q("IconStyle"))
        ic_h = ET.SubElement(is_h, Q("Icon"))
        ET.SubElement(ic_h, Q("href")).text = href
        ls_h = ET.SubElement(st_h, Q("LabelStyle"))
        ET.SubElement(ls_h, Q("scale")).text = "1"
        ET.SubElement(ls_h, Q("color")).text = "ffffffff"

        sm = ET.Element(Q("StyleMap"), {"id": sm_id})
        p1 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p1, Q("key")).text = "normal"
        ET.SubElement(p1, Q("styleUrl")).text = f"#{sm_id}_normal"
        p2 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p2, Q("key")).text = "highlight"
        ET.SubElement(p2, Q("styleUrl")).text = f"#{sm_id}_highlight"

        insert_before_first_folder(st_n)
        insert_before_first_folder(st_h)
        insert_before_first_folder(sm)
        key_to_smid[(href, hide_flag)] = sm_id

    for pm, href, hide_flag in pm_info:
        smid = key_to_smid.get((href, hide_flag))
        if not smid:
            continue
        for existing in pm.findall(Q("styleUrl")):
            pm.remove(existing)
        for inline_style in pm.findall(Q("Style")):
            pm.remove(inline_style)
        ET.SubElement(pm, Q("styleUrl")).text = f"#{smid}"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

# -------------------------
# UI: load xlsx
# -------------------------
uploaded_xlsx = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])
if not uploaded_xlsx:
    st.stop()

try:
    df_dict = pd.read_excel(uploaded_xlsx, sheet_name=None)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

normalized = {k.strip().upper(): v for k, v in df_dict.items()}

def get_sheet(*names):
    for n in names:
        if not n:
            continue
        key = n.strip().upper()
        df = normalized.get(key)
        if df is not None and not df.empty:
            return df
    return None

df_agms = get_sheet("AGMS", "AGM")
df_access = get_sheet("ACCESS")
df_center = get_sheet("CENTERLINE")
df_notes = get_sheet("NOTES")

tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])
with tab1:
    st.subheader("AGMs")
    st.dataframe(df_agms if df_agms is not None else pd.DataFrame())
with tab2:
    st.subheader("Access")
    st.dataframe(df_access if df_access is not None else pd.DataFrame())
with tab3:
    st.subheader("Centerline")
    st.dataframe(df_center if df_center is not None else pd.DataFrame())
with tab4:
    st.subheader("Notes")
    st.dataframe(df_notes if df_notes is not None else pd.DataFrame())

# -------------------------
# Generate KMZ
# -------------------------
if st.button("Generate KMZ"):
    kml = simplekml.Kml()

    # Gibson logo on every output
    add_gibson_logo_overlay(kml)

    # Notes hide flags
    notes_flags_by_name = {}
    hide_col = None
    if df_notes is not None:
        for c in df_notes.columns:
            if str(c).strip().lower() == "hidenameuntilmouseover":
                hide_col = c
                break
        for _, row in df_notes.iterrows():
            nm = str(safe_str(row.get("Name")) or "").strip()
            hide_flag = True
            if hide_col:
                v = row.get(hide_col)
                if pd.notna(v) and str(v).strip().lower() in ("0", "false", "no", "n", "f"):
                    hide_flag = False
                elif pd.notna(v) and str(v).strip().lower() in ("1", "true", "yes", "y", "t"):
                    hide_flag = True
            notes_flags_by_name[nm] = hide_flag

    # AGMs
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_agm_point(folder, row)

    # Access (keeps LineStringColor)
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_lines_with_autosplit(folder, df_access, color_col="LineStringColor", split_jump_m=5000.0)
        if not created:
            for _, row in df_access.iterrows():
                add_access_point(folder, row)

    # Centerline (split on big jumps so it won't connect distant blocks)
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        created = add_lines_with_autosplit(folder, df_center, color_col="LineStringColor", split_jump_m=5000.0)
        if not created:
            for _, row in df_center.iterrows():
                add_access_point(folder, row)

    # Notes
    if df_notes is not None:
        folder = kml.newfolder(name="Notes")
        for _, row in df_notes.iterrows():
            add_note_point(folder, row)

    # Build + inject hover styles for Notes only
    try:
        raw_kml = kml.kml().encode("utf-8")
        modified_kml = inject_hover_stylemaps_for_notes_with_flags(
            raw_kml,
            notes_flags_by_name=notes_flags_by_name,
            notes_folder_name="Notes"
        )
    except Exception as e:
        st.error(f"Failed to build or modify KML: {e}")
        st.stop()

    # Package KMZ
    kmz_bytes = io.BytesIO()
    try:
        with zipfile.ZipFile(kmz_bytes, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("doc.kml", modified_kml)
            zf.writestr("logo.png", base64.b64decode(GIBSON_LOGO_B64))
    except Exception as e:
        st.error(f"Failed to build KMZ: {e}")
        st.stop()

    st.download_button(
        label="Download KMZ",
        data=kmz_bytes.getvalue(),
        file_name="KMZ_Generator_Output.kmz",
        mime="application/vnd.google-earth.kmz"
    )
    st.success("KMZ generated successfully.")
