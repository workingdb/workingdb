SELECT
    newPurch.NIFCO_PART_NUMBER AS partNumber,
    'empty' AS Part_Picture,
    newPurch.PART_DESCRIPTION,
    newPurch.SIFTYPE,
    newPurch.DESIGN_RESPONSIBILITY,
    newPurch.COMPONENT_OR_FINISHED_GOOD AS GOOD_TYPE,
    newPurch.CUSTOMER_PART_NUM AS CUST_PN,
    newPurch.CREATION_DATE AS createdDate,
    Trim(
        SUBSTR (
            O_E_MAKER_MODEL,
            1,
            INSTR (O_E_MAKER_MODEL, ' - ')
        )
    ) AS MODEL_NAME,
    Trim(
        SUBSTR (
            O_E_MAKER_MODEL,
            INSTR (O_E_MAKER_MODEL, ' - ') + 2
        )
    ) AS OE_Maker,
    newPurch.MODELYR AS modelYear,
    newPurch.NUMBER_OF_PIECES_PER_VEHICLE AS PPV,
    newPurch.EST_MONTLY_VOLUME AS monthlyVol,
    newPurch.PIECE_PRICE AS piecePrice,
    newPurch.INTERNAL_PART_COST AS pieceCost,
    CAT.CLASS_CODE,
    CAT.SUB_CLASS_CODE,
    CAT.BUSINESS_CODE,
    CAT.FOCUS_AREA_CODE
FROM
    APPS.Q_SIF_NEW_PURCHASING_PART_V newPurch
    LEFT JOIN (
        SELECT
            sysItems.SEGMENT1 AS PN,
            catVL.SEGMENT1 as CLASS_CODE,
            catVL.SEGMENT2 as SUB_CLASS_CODE,
            catVL.SEGMENT3 AS BUSINESS_CODE,
            catVL.SEGMENT4 AS FOCUS_AREA_CODE,
            catVL.STRUCTURE_ID
        FROM
            (
                INV.MTL_ITEM_CATEGORIES ItemCat
                LEFT JOIN APPS.MTL_CATEGORIES_VL catVL ON ItemCat.CATEGORY_ID = catVL.CATEGORY_ID
            )
            INNER JOIN INV.MTL_SYSTEM_ITEMS_B sysItems ON ItemCat.INVENTORY_ITEM_ID = sysItems.INVENTORY_ITEM_ID
        GROUP BY
            sysItems.SEGMENT1,
            catVL.SEGMENT1,
            catVL.SEGMENT2,
            catVL.SEGMENT3,
            catVL.SEGMENT4,
            catVL.STRUCTURE_ID
        HAVING
            catVL.STRUCTURE_ID = 50569
    ) CAT ON newPurch.NIFCO_PART_NUMBER = CAT.PN
UNION SELECT
    newAss.NIFCO_PART_NUMBER AS partNumber,
    'empty' AS Part_Picture,
    newAss.PART_DESCRIPTION,
    newAss.SIFTYPE,
    newAss.DESIGN_RESPONSIBILITY,
    newAss.FINISHED_GOOD_OR_SUBASSEMBLY AS GOOD_TYPE,
    newAss.CUSTOMER_PART_NUM AS CUST_PN,
    newAss.CREATION_DATE AS createdDate,
    Trim(
        SUBSTR (
            O_E_MAKER_MODEL,
            1,
            INSTR (O_E_MAKER_MODEL, ' - ')
        )
    ) AS MODEL_NAME,
    Trim(
        SUBSTR (
            O_E_MAKER_MODEL,
            INSTR (O_E_MAKER_MODEL, ' - ') + 2
        )
    ) AS OE_Maker,
    newAss.MODELYR AS modelYear,
    newAss.NUMBER_OF_PIECES_PER_VEHICLE AS PPV,
    newAss.EST_MONTLY_VOLUME AS monthlyVol,
    newAss.PIECE_PRICE AS piecePrice,
    newAss.INTERNAL_PART_COST AS pieceCost,
    CAT.CLASS_CODE,
    CAT.SUB_CLASS_CODE,
    CAT.BUSINESS_CODE,
    CAT.FOCUS_AREA_CODE
FROM
    APPS.Q_SIF_NEW_ASSEMBLED_PART_V newAss
    LEFT JOIN (
        SELECT
            sysItems.SEGMENT1 AS PN,
            catVL.SEGMENT1 as CLASS_CODE,
            catVL.SEGMENT2 as SUB_CLASS_CODE,
            catVL.SEGMENT3 AS BUSINESS_CODE,
            catVL.SEGMENT4 AS FOCUS_AREA_CODE,
            catVL.STRUCTURE_ID
        FROM
            (
                INV.MTL_ITEM_CATEGORIES ItemCat
                LEFT JOIN APPS.MTL_CATEGORIES_VL catVL ON ItemCat.CATEGORY_ID = catVL.CATEGORY_ID
            )
            INNER JOIN INV.MTL_SYSTEM_ITEMS_B sysItems ON ItemCat.INVENTORY_ITEM_ID = sysItems.INVENTORY_ITEM_ID
        GROUP BY
            sysItems.SEGMENT1,
            catVL.SEGMENT1,
            catVL.SEGMENT2,
            catVL.SEGMENT3,
            catVL.SEGMENT4,
            catVL.STRUCTURE_ID
        HAVING
            catVL.STRUCTURE_ID = 50569
    ) CAT ON newAss.NIFCO_PART_NUMBER = CAT.PN
UNION SELECT
    newMold.NIFCO_PART_NUMBER AS partNumber,
    'empty' AS Part_Picture,
    newMold.PART_DESCRIPTION,
    newMold.SIFTYPE,
    newMold.DESIGN_RESPONSIBILITY,
    newMold.COMPONENT_OR_FINISHED_GOOD AS GOOD_TYPE,
    newMold.NEWCUSTOMERPARTNUM AS CUST_PN,
    newMold.CREATION_DATE AS createdDate,
    Trim(
        SUBSTR (
            O_E_MAKER_MODEL,
            1,
            INSTR (O_E_MAKER_MODEL, ' - ')
        )
    ) AS MODEL_NAME,
    Trim(
        SUBSTR (
            O_E_MAKER_MODEL,
            INSTR (O_E_MAKER_MODEL, ' - ') + 2
        )
    ) AS OE_Maker,
    newMold.MODELYR AS modelYear,
    newMold.NUMBER_OF_PIECES_PER_VEHICLE AS PPV,
    newMold.EST_MONTLY_VOLUME AS monthlyVol,
    newMold.PIECE_PRICE AS piecePrice,
    newMold.INTERNAL_PART_COST AS pieceCost,
    CAT.CLASS_CODE,
    CAT.SUB_CLASS_CODE,
    CAT.BUSINESS_CODE,
    CAT.FOCUS_AREA_CODE
FROM
    APPS.Q_SIF_NEW_MOLDED_PART_V newMold
    LEFT JOIN (
        SELECT
            sysItems.SEGMENT1 AS PN,
            catVL.SEGMENT1 as CLASS_CODE,
            catVL.SEGMENT2 as SUB_CLASS_CODE,
            catVL.SEGMENT3 AS BUSINESS_CODE,
            catVL.SEGMENT4 AS FOCUS_AREA_CODE,
            catVL.STRUCTURE_ID
        FROM
            (
                INV.MTL_ITEM_CATEGORIES ItemCat
                LEFT JOIN APPS.MTL_CATEGORIES_VL catVL ON ItemCat.CATEGORY_ID = catVL.CATEGORY_ID
            )
            INNER JOIN INV.MTL_SYSTEM_ITEMS_B sysItems ON ItemCat.INVENTORY_ITEM_ID = sysItems.INVENTORY_ITEM_ID
        GROUP BY
            sysItems.SEGMENT1,
            catVL.SEGMENT1,
            catVL.SEGMENT2,
            catVL.SEGMENT3,
            catVL.SEGMENT4,
            catVL.STRUCTURE_ID
        HAVING
            catVL.STRUCTURE_ID = 50569
    ) CAT ON newMold.NIFCO_PART_NUMBER = CAT.PN
