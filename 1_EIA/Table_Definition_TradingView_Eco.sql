CREATE TABLE IF NOT EXISTS public.tradingview_eco
(
    id character varying(10) COLLATE pg_catalog."default" NOT NULL,
    event_date date NOT NULL,
    event_time time without time zone NOT NULL,
    event_name character varying(50) COLLATE pg_catalog."default" NOT NULL,
    actual numeric(20,4),
    forecast numeric(20,4),
    previous numeric(20,4),
    actual_forecast numeric(10,4),
    actual_previous numeric(10,4),
    importance smallint,
    report_period character varying(25) COLLATE pg_catalog."default",
    unit character varying(25) COLLATE pg_catalog."default",
    scale character(1) COLLATE pg_catalog."default",
    update_date date,
    update_time time without time zone,
    update_news boolean,
    update_pre boolean,
    futures_relation_1 character varying(10) COLLATE pg_catalog."default",
    futures_relation_2 character varying(10) COLLATE pg_catalog."default",
    specific_relation_1 character varying(10) COLLATE pg_catalog."default",
    specific_relation_2 character varying(10) COLLATE pg_catalog."default",
    specific_relation_3 character varying(10) COLLATE pg_catalog."default",
    specific_relation_4 character varying(10) COLLATE pg_catalog."default",
    specific_relation_5 character varying(10) COLLATE pg_catalog."default",
    specific_relation_6 character varying(10) COLLATE pg_catalog."default",
    specific_relation_7 character varying(10) COLLATE pg_catalog."default",
    url character varying(375) COLLATE pg_catalog."default" NOT NULL,
    ignored boolean,
    important boolean,
    CONSTRAINT tradingview_eco_pkey PRIMARY KEY (id, event_date, event_time, event_name)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS public.tradingview_eco
    OWNER to postgres;