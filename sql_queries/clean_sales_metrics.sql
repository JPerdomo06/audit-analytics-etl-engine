SQL

-- Title: Data Cleaning & Revenue Leakage Audit Query
-- Description: Groups sales by product, calculates descriptive statistics, 
--              and excludes transactions linked to credit notes to ensure net revenue integrity.
-- Database Dialect: Oracle SQL / Standard SQL

SELECT
    s.sku_id                        AS sku,
    s.product_description           AS product_description,
    AVG(s.list_price)               AS avg_list_price,
    AVG(s.selling_price)            AS avg_selling_price,
    SUM(s.units_sold)               AS total_units_sold,
    AVG(s.discounts_applied)        AS avg_discounts,
    AVG(s.taxable_sales_tier_1)     AS avg_taxable_sales_15pct,
    AVG(s.taxable_sales_tier_2)     AS avg_taxable_sales_18pct,
    AVG(s.exempt_sales)             AS avg_exempt_sales,
    AVG(s.vat_collected_tier_1)     AS avg_vat_15pct,
    AVG(s.vat_collected_tier_2)     AS avg_vat_18pct,
    AVG(s.total_item_revenue)       AS avg_total_item_revenue
FROM sales_ledger s
LEFT JOIN (
    -- Subquery to consolidate all unique document references linked to credit notes
    SELECT DISTINCT UPPER(TRIM(invoice_id)) AS doc_reference
    FROM credit_notes
    WHERE invoice_id IS NOT NULL AND TRIM(invoice_id) != ''

    UNION
    
    SELECT DISTINCT UPPER(TRIM(origin_doc_id)) AS doc_reference
    FROM credit_notes
    WHERE origin_doc_id IS NOT NULL AND TRIM(origin_doc_id) != ''
) cn
    ON UPPER(TRIM(s.invoice_correlative)) = cn.doc_reference
WHERE
    s.invoice_correlative IS NOT NULL
    AND TRIM(s.invoice_correlative) != ''
    AND cn.doc_reference IS NULL -- Strictly excludes any invoice matched with a credit note
GROUP BY
    s.sku_id, 
    s.product_description;
