SELECT  syscolumns.name AS Name ,
        systypes.name AS DataType ,
        syscolumns.length AS Length ,
        CASE syscolumns.isnullable
          WHEN 1 THEN 'Y'
          ELSE 'N'
        END AS 'Nullable' ,
        ISNULL(sys.extended_properties.value, '') AS Description
FROM    syscolumns
        JOIN systypes ON syscolumns.xtype = systypes.xtype
                         AND systypes.name <> 'sysname'
        LEFT OUTER JOIN sys.extended_properties ON ( sys.extended_properties.minor_id = syscolumns.colid
                                                     AND sys.extended_properties.major_id = syscolumns.id
                                                   )
WHERE   syscolumns.id IN ( SELECT   id
                           FROM     sysobjects
                           WHERE    name = 'RMATrackingNumberTransaction' )
ORDER BY NAME
 