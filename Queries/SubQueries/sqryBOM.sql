SELECT
    sysItems.SEGMENT1 AS Assy,
    sysItems1.SEGMENT1 AS Compt,
    bomInv.IMPLEMENTATION_DATE,
    bomMat.ORGANIZATION_ID,
    bomInv.DISABLE_DATE,
    Round(COMPONENT_QUANTITY, 5) as Qty,
    Round(1 / COMPONENT_QUANTITY, 5) as Inverse_Qty,
    sysItems.DESCRIPTION as assyDescription,
    sysItems.INVENTORY_ITEM_STATUS_CODE as assyStatus,
    sysItems1.DESCRIPTION as compDescription,
    sysItems1.INVENTORY_ITEM_STATUS_CODE as compStatus,
    COMPONENT_ITEM_ID,
    ASSEMBLY_ITEM_ID
FROM
    (
        (
            APPS.BOM_BILL_OF_MATERIALS bomMat
            INNER JOIN APPS.BOM_INVENTORY_COMPONENTS bomInv ON bomMat.COMMON_BILL_SEQUENCE_ID = bomInv.BILL_SEQUENCE_ID
        )
        INNER JOIN INV.MTL_SYSTEM_ITEMS_B sysItems ON bomMat.ASSEMBLY_ITEM_ID = sysItems.INVENTORY_ITEM_ID
    )
    INNER JOIN INV.MTL_SYSTEM_ITEMS_B sysItems1 ON bomInv.COMPONENT_ITEM_ID = sysItems1.INVENTORY_ITEM_ID
                            WHERE (sysItems.SEGMENT1 = '26587' AND bomInv.DISABLE_DATE Is Null);
