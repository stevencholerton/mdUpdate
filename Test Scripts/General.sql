SELECT * FROM routingmethods 
WHERE partid = 'SH0001.CGR.CGR'
AND routingmethods.methodid = 'ALT'

SELECT * FROM routingoperations r
LEFT JOIN routingmethods rm
    ON rm.partid = r.partid
	AND rm.methodid = r.methodid
WHERE rm.partid = 'SH0001.CGR.CGR'

SELECT * FROM routingoperations r
WHERE r.partid = 'SH0001.CGR.CGR'

UPDATE routingmethods
SET description = 'TEST010'
FROM routingoperations ro
WHERE routingmethods.partid = 'SH0001.CGR.CGR'
AND routingmethods.methodid = 'ALT'
AND ro.partid = 'SH0001.CGR.CGR'
AND ro.methodid = 'ALT'
AND ro.operationnumber = '30'