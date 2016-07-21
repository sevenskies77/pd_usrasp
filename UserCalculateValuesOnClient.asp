<%
'Autocalculate field values on the client side
'ExtGUID - data source GUID
'FieldToCalculate - field name which has to be calculated
'Formula - VBScript-formula valid on the client side
'
'Possible functions:
'UserCalculateValuesOnClient ExtGUID, FieldToCalculate, Formula
'UserCalculateValuesOnClientShow ExtGUID, FieldToCalculate, Formula, bShow
'
'Use the expression forma.FieldName.value to provide the value of FieldName field
'Example:
'
'UserCalculateValuesOnClient "{60586E06-E3F3-4214-9877-8422A001C165}", "MoneyField", "forma.IntegerField.value/100*20"
'UserCalculateValuesOnClient "{60586E06-E3F3-4214-9877-8422A001C165}", "MoneyField", "CCur(IIF(Trim(forma.IntegerField.value)<>"""",forma.IntegerField.value,""0""))/100*20"
'UserCalculateValuesOnClient "{60586E06-E3F3-4214-9877-8422A001C165}", "MoneyField", "IIF(Trim(forma.IntegerField.value)<>"""",forma.IntegerField.value,""0"")/100*20"
'UserCalculateValuesOnClient "{60586E06-E3F3-4214-9877-8422A001C167}", "DateField", "Now"

'UserCalculateValuesOnClient "{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}", ToEngName("Сумма"), "forma."+ToEngName("Цена (тариф за ед изм) руб")+".value*CCur(forma."+ToEngName("Количество")+".value)"
'UserCalculateValuesOnClient "{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}", ToEngName("Налоговая ставка"), "15"
'UserCalculateValuesOnClient "{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}", ToEngName("Сумма налога, руб"), "forma."+ToEngName("Сумма")+".value/100*CCur(forma."+ToEngName("Налоговая ставка")+".value)"
'UserCalculateValuesOnClient "{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}", ToEngName("Стоимость товаров (работ, услуг), всего с учетом налога, руб"), "forma."+ToEngName("Сумма")+".value+CCur(forma."+ToEngName("Сумма налога, руб")+".value)"

'UserCalculateValuesOnClient "{C3DB86C0-0F73-FFBA-DF8D-9F57ED31707A}", "Amount", "CCur(forma.Qty.value)*CCur(forma.Price.value)"
'UserCalculateValuesOnClient "{C3DB86C0-0F73-FFBA-DF8D-9F57ED31707A}", "Tax", "CCur(forma.Amount.value)/100*15"
'UserCalculateValuesOnClient "{C3DB86C0-0F73-FFBA-DF8D-9F57ED31707A}", "AmountPlusTax", "CCur(forma.Amount.value)+CCur(forma.Tax.value)"

'UserCalculateValuesOnClient "{5EB5017F-334A-6303-BC09-BADCECCF8657}", "BillToName", "IIF(forma.BillingSameAsShipping.value=""YES"",forma.ShipToName.value, forma.BillToName.value)"
'UserCalculateValuesOnClient "{5EB5017F-334A-6303-BC09-BADCECCF8657}", "BillToAddress", "IIF(forma.BillingSameAsShipping.value=""YES"",forma.ShipToAddress.value, forma.BillToAddress.value)"
'UserCalculateValuesOnClient "{5EB5017F-334A-6303-BC09-BADCECCF8657}", "BillToCity", "IIF(forma.BillingSameAsShipping.value=""YES"",forma.ShipToCity.value, forma.BillToCity.value)"
'UserCalculateValuesOnClient "{5EB5017F-334A-6303-BC09-BADCECCF8657}", "BillToState", "IIF(forma.BillingSameAsShipping.value=""YES"",forma.ShipToState.value, forma.BillToState.value)"
'UserCalculateValuesOnClient "{5EB5017F-334A-6303-BC09-BADCECCF8657}", "BillToZipCode", "IIF(forma.BillingSameAsShipping.value=""YES"",forma.ShipToZipCode.value, forma.BillToZipCode.value)"
'UserCalculateValuesOnClient "{5EB5017F-334A-6303-BC09-BADCECCF8657}", "BillToPhone", "IIF(forma.BillingSameAsShipping.value=""YES"",forma.ShipToPhone.value, forma.BillToPhone.value)"

'UserSetDefaultValuesOnClient "{6336D704-AFF9-2274-64F4-AA7D3E8F3AE2}", ToEngName("Номер грузовой таможенной декларации"), """0000"""

%>
