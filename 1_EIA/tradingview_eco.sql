SELECT event_date,
event_time,
event_name,
actual,
forecast,
previous
FROM tradingview_eco
WHERE event_date >= CURRENT_DATE - 100 AND
event_date <= CURRENT_DATE AND 
( event_name like 'EIA Wkly Crude Stk%' OR 
 event_name like 'EIA Wkly Dist. Stk%' OR 
 event_name like 'EIA Wkly Gsln Stk%' OR 
 event_name like 'API Wkly dist. Stk%' OR
 event_name like 'API Wkly gsln stk%' )
 ORDER BY event_date, event_time, event_name;
