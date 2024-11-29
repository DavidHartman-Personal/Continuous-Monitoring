-- sequence for primary key column
CREATE SEQUENCE public.poam_header_id_seq
    INCREMENT 1
    START 2
    MINVALUE 1
    MAXVALUE 2147483647
    CACHE 1;

ALTER SEQUENCE public.poam_header_id_seq
    OWNER TO postgres;

DROP  TABLE public.poam_header;

CREATE TABLE public.poam_header
(
    id                            integer               NOT NULL DEFAULT nextval('poam_header_id_seq'::regclass),
    poam_id                       character varying(30) NOT NULL,
--    controls                      character varying(30),
    weakness_name                 character varying(250),
    weakness_description          character varying(1000),
--    weakness_detector_source      character varying(50),
    weakness_source_identifier    character varying(100),
    asset_identifier              character varying(150),
--     point_of_contact              character varying(150),
--     resources_required            character varying(1000),
--     overall_remediation_plan      character varying(1000),
--     original_detection_date       timestamp with time zone,
--     scheduled_completion_date     timestamp with time zone,
--     planned_milestones            character varying(1000),
--     milestone_changes             character varying(1000),
--     status_date                   timestamp with time zone,
--     vendor_dependency             boolean               NOT NULL DEFAULT FALSE,
--     last_vendor_check_in_date     timestamp with time zone,
--     vendor_dependent_product_name character varying(150),
--     original_risk_rating          character varying(150),
--     adjusted_risk_rating          character varying(150),
--     risk_adjustment               boolean               NOT NULL DEFAULT FALSE,
--     false_positive                boolean               NOT NULL DEFAULT FALSE,
--     operational_requirement       boolean               NOT NULL DEFAULT FALSE,
--     deviation_rationale           character varying(1000),
--     supporting_documents          character varying(500),
--     comments                      character varying(1000),
--     auto_approve                  character varying(150),
    month_year_added              character varying(10),
    create_date                   timestamp with time zone NOT NULL DEFAULT CURRENT_TIMESTAMP,
    update_date                   timestamp with time zone,
    CONSTRAINT poam_header_pkey PRIMARY KEY (ID)
--    CONSTRAINT poam_header_poam_id UNIQUE (POAM_ID)
)

TABLESPACE pg_default;

ALTER TABLE public.poam_header
    OWNER to postgres;

COMMENT ON COLUMN public.poam_header.month_year_added
    IS 'Month and year that POAM was added';


-- DROP  TABLE public.poam_header_notes;

CREATE TABLE public.poam_header_notes
(
    id                            integer               NOT NULL DEFAULT nextval('poam_header_id_seq'::regclass),
    poam_id                       character varying(30) NOT NULL,
    note_date                     timestamp with time zone NOT NULL DEFAULT CURRENT_TIMESTAMP,
    note                          character varying(1000),
    create_date                   timestamp with time zone NOT NULL DEFAULT CURRENT_TIMESTAMP,
    update_date                   timestamp with time zone,
    CONSTRAINT poam_header_pkey PRIMARY KEY (id)
)

TABLESPACE pg_default;

ALTER TABLE public.poam_header_notes
    OWNER to postgres;


-- Table: public.auth_user

-- DROP TABLE public.auth_user;

-- CREATE TABLE public.auth_user
-- (
--     id integer NOT NULL DEFAULT nextval('auth_user_id_seq'::regclass),
--     password character varying(128) COLLATE pg_catalog."default" NOT NULL,
--     last_login timestamp with time zone,
--     is_superuser boolean NOT NULL,
--     username character varying(150) COLLATE pg_catalog."default" NOT NULL,
--     first_name character varying(30) COLLATE pg_catalog."default" NOT NULL,
--     last_name character varying(150) COLLATE pg_catalog."default" NOT NULL,
--     email character varying(254) COLLATE pg_catalog."default" NOT NULL,
--     is_staff boolean NOT NULL,
--     is_active boolean NOT NULL,
--     date_joined timestamp with time zone NOT NULL,
--     CONSTRAINT auth_user_pkey PRIMARY KEY (id),
--     CONSTRAINT auth_user_username_key UNIQUE (username)
--
-- )
--
-- TABLESPACE pg_default;
--
-- ALTER TABLE public.auth_user
--     OWNER to postgres;
--
-- -- Index: auth_user_username_6821ab7c_like
--
-- -- DROP INDEX public.auth_user_username_6821ab7c_like;
--
-- CREATE INDEX auth_user_username_6821ab7c_like
--     ON public.auth_user USING btree
--     (username COLLATE pg_catalog."default" varchar_pattern_ops)
--     TABLESPACE pg_default;
--
--
-- -- SEQUENCE: public.auth_user_id_seq
--
-- -- DROP SEQUENCE public.auth_user_id_seq;
--

