-- Enum Type DDL

-- 商品区分
DROP TYPE IF EXISTS product_kbn;
CREATE TYPE product_kbn AS enum (
  'NO_CHECKED',
  'NORMAL',
  'CHANGED',
  'STOPPED'
);

-- 部署
DROP TYPE IF EXISTS department;
CREATE TYPE department AS enum (
  'RECEIVING',
  'ORDERING',
  'PRODUCTS',
  'OVERSEAS',
  'ACCOUNTING'
);

-- 取引先区分
DROP TYPE IF EXISTS customer_type;
CREATE TYPE customer_type AS enum (
  'LOYAL',
  'STANDARD',
  'TRIAL',
  'NOT_TRADE'
);

