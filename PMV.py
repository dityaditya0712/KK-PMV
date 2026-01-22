# Lib
import os
import sys
import glob
import re
import numpy as np
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

# Transformasi Form Syariah ke Konvensional
## Laba Rugi
def transform_laba_rugi(df_sya):
    """
    Transformasi df_sya untuk laporan Laba Rugi.
    """
    df_sya["BBN_BG_HSL_TS_PNDNN_YNG_DTRM_BRDSRKN_PRNSP_SYRH_TTL"] = (
        df_sya["BBN_BNG_DR_PNJMNPMBYN_YNG_DTRM_TTL"]
        + df_sya["BBN_BNG_DR_SRT_BRHRG_YNG_DTRBTKN_TTL"]
    ).fillna(0)

    df_sya["BBN_BG_HSL_TS_PNDNN_YNG_DTRM_BRDSRKN_PRNSP_SYRH_NDNSN_RPH"] = (
        df_sya["BBN_BNG_DR_PNJMNPMBYN_YNG_DTRM_NDNSN_RPH"]
        + df_sya["BBN_BNG_DR_SRT_BRHRG_YNG_DTRBTKN_NDNSN_RPH"]
    ).fillna(0)

    df_sya["BBN_BG_HSL_TS_PNDNN_YNG_DTRM_BRDSRKN_PRNSP_SYRH_MT_NG_SNG"] = (
        df_sya["BBN_BNG_DR_PNJMNPMBYN_YNG_DTRM_MT_NG_SNG"]
        + df_sya["BBN_BNG_DR_SRT_BRHRG_YNG_DTRBTKN_MT_NG_SNG"]
    ).fillna(0)

    return df_sya
## Arus Kas
def transform_arus_kas(df_sya):
    """
    Transformasi kolom Arus Kas untuk prinsip Syariah.
    Menghitung total, rupiah, dan mata uang asing untuk masuk dan keluar kas.
    """

    # -----------------------------
    # Arus Kas Masuk Dari Kegiatan Investasi Berdasarkan Prinsip Syariah
    # -----------------------------
    df_sya["RS_KS_MSK_DR_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH_NDNSN_RPH"] = (
        df_sya["RS_KS_MSK_DR_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_NDNSN_RPH"]
        + df_sya["RS_KS_MSK_DR_KGTN_PNYRTN_MLL_PMBLN_SKK_T_BLGS_SYRH_KNVRS_NDNSN_RPH"]
        + df_sya["RS_KS_MSK_DR_KGTN_PMBYN_MLL_PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_THP_RNTSN_WL_STRTP_DNT_PNGMBNGN_SH_NDNSN_RPH"]
        + df_sya["RS_KS_MSK_DR_KGTN_PMBYN_BRDSRKN_PRNSP_BG_HSL_NDNSN_RPH"]
    ).fillna(0)

    df_sya["RS_KS_MSK_DR_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH_MT_NG_SNG"] = (
        df_sya["RS_KS_MSK_DR_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_MT_NG_SNG"]
        + df_sya["RS_KS_MSK_DR_KGTN_PNYRTN_MLL_PMBLN_SKK_T_BLGS_SYRH_KNVRS_MT_NG_SNG"]
        + df_sya["RS_KS_MSK_DR_KGTN_PMBYN_MLL_PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_THP_RNTSN_WL_STRTP_DNT_PNGMBNGN_SH_MT_NG_SNG"]
        + df_sya["RS_KS_MSK_DR_KGTN_PMBYN_BRDSRKN_PRNSP_BG_HSL_MT_NG_SNG"]
    ).fillna(0)

    df_sya["RS_KS_MSK_DR_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH_TTL"] = (
        df_sya["RS_KS_MSK_DR_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_TTL"]
        + df_sya["RS_KS_MSK_DR_KGTN_PNYRTN_MLL_PMBLN_SKK_T_BLGS_SYRH_KNVRS_TTL"]
        + df_sya["RS_KS_MSK_DR_KGTN_PMBYN_MLL_PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_THP_RNTSN_WL_STRTP_DNT_PNGMBNGN_SH_TTL"]
        + df_sya["RS_KS_MSK_DR_KGTN_PMBYN_BRDSRKN_PRNSP_BG_HSL_TTL"]
    ).fillna(0)

    # -----------------------------
    # Arus Kas Masuk Dari Kegiatan Jasa Berbasis Fee
    # -----------------------------
    df_sya["RS_KS_MSK_DR_KGTN_JS_BRBSS_F_NDNSN_RPH"] = df_sya["RS_KS_MSK_DR_KGTN_JS_BRBSS_MBL_HSL_NDNSN_RPH"].fillna(0)
    df_sya["RS_KS_MSK_DR_KGTN_JS_BRBSS_F_MT_NG_SNG"] = df_sya["RS_KS_MSK_DR_KGTN_JS_BRBSS_MBL_HSL_MT_NG_SNG"].fillna(0)
    df_sya["RS_KS_MSK_DR_KGTN_JS_BRBSS_F_TTL"] = df_sya["RS_KS_MSK_DR_KGTN_JS_BRBSS_MBL_HSL_TTL"].fillna(0)

    # -----------------------------
    # Arus Kas Keluar Untuk Kegiatan Investasi Berdasarkan Prinsip Syariah
    # -----------------------------
    df_sya["RS_KS_KLR_NTK_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH__NDNSN_RPH"] = (
        df_sya["RS_KS_KLR_NTK_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_NDNSN_RPH"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBLN_SKK_T_BLGS_SYRH_KNVRS_NDNSN_RPH"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_THP_RNTSN_WL_STRTP_DNT_PNGMBNGN_SH_NDNSN_RPH"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBYN_BRDSRKN_PRNSP_BG_HSL_NDNSN_RPH"]
    ).fillna(0)

    df_sya["RS_KS_KLR_NTK_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH__MT_NG_SNG"] = (
        df_sya["RS_KS_KLR_NTK_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_MT_NG_SNG"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBLN_SKK_T_BLGS_SYRH_KNVRS_MT_NG_SNG"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_THP_RNTSN_WL_STRTP_DNT_PNGMBNGN_SH_MT_NG_SNG"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBYN_BRDSRKN_PRNSP_BG_HSL_MT_NG_SNG"]
    ).fillna(0)

    df_sya["RS_KS_KLR_NTK_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH__TTL"] = (
        df_sya["RS_KS_KLR_NTK_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_TTL"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBLN_SKK_T_BLGS_SYRH_KNVRS_TTL"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_THP_RNTSN_WL_STRTP_DNT_PNGMBNGN_SH_TTL"]
        + df_sya["RS_KS_KLR_NTK_KGTN_PMBYN_BRDSRKN_PRNSP_BG_HSL_TTL"]
    ).fillna(0)

    # -----------------------------
    # Arus Kas Keluar Untuk Kegiatan Jasa Berbasis Fee
    # -----------------------------
    df_sya["RS_KS_KLR_NTK_KGTN_JS_BRBSS_F_NDNSN_RPH"] = df_sya["RS_KS_KLR_DR_KGTN_JS_BRBSS_MBL_HSL_NDNSN_RPH"].fillna(0)
    df_sya["RS_KS_KLR_NTK_KGTN_JS_BRBSS_F_MT_NG_SNG"] = df_sya["RS_KS_KLR_DR_KGTN_JS_BRBSS_MBL_HSL_MT_NG_SNG"].fillna(0)
    df_sya["RS_KS_KLR_NTK_KGTN_JS_BRBSS_F_TTL"] = df_sya["RS_KS_KLR_DR_KGTN_JS_BRBSS_MBL_HSL_TTL"].fillna(0)

    # -----------------------------
    # Arus Kas Keluar Untuk Pendanaan Berdasarkan Prinsip Syariah
    # -----------------------------
    df_sya["RS_KS_KLR_NTK_PNDNN_BRDSRKN_PRNSP_SYRH_NDNSN_RPH"] = (
        df_sya["RS_KS_KLR_NTK_PMBYRN_PKK_PNJMN_DN_SRT_BRHRG_YNG_DTRBTKN_NDNSN_RPH"]
        + df_sya["RS_KS_KLR_NTK_PNRKN_KMBL_MDL_PRSHN_TRSRY_STCK_NDNSN_RPH"]
        + df_sya["RS_KS_KLR_NTK_PMBYRN_DVDN_NDNSN_RPH"]
    ).fillna(0)

    df_sya["RS_KS_KLR_NTK_PNDNN_BRDSRKN_PRNSP_SYRH_MT_NG_SNG"] = (
        df_sya["RS_KS_KLR_NTK_PMBYRN_PKK_PNJMN_DN_SRT_BRHRG_YNG_DTRBTKN_MT_NG_SNG"]
        + df_sya["RS_KS_KLR_NTK_PNRKN_KMBL_MDL_PRSHN_TRSRY_STCK_MT_NG_SNG"]
        + df_sya["RS_KS_KLR_NTK_PMBYRN_DVDN_MT_NG_SNG"]
    ).fillna(0)

    df_sya["RS_KS_KLR_NTK_PNDNN_BRDSRKN_PRNSP_SYRH_TTL"] = (
        df_sya["RS_KS_KLR_NTK_PMBYRN_PKK_PNJMN_DN_SRT_BRHRG_YNG_DTRBTKN_TTL"]
        + df_sya["RS_KS_KLR_NTK_PNRKN_KMBL_MDL_PRSHN_TRSRY_STCK_TTL"]
        + df_sya["RS_KS_KLR_NTK_PMBYRN_DVDN_TTL"]
    ).fillna(0)

    return df_sya

## Rekening Administratif
def transform_rekening_administratif(df_sya):
    # Penyaluran Pembiayaan Bersama
    df_sya["SLD_TSTNDNG_PRNCPLS_PNYLRN_PMBYN_BRSM_NDNSN_RPH"] = df_sya["KGTN_PMBYN_PNRSN_CHNLLNG_TTL"].fillna(0)
    df_sya["SLD_TSTNDNG_PRNCPLS_PNYLRN_PMBYN_BRSM_MT_NG_SNG"] = df_sya["KGTN_PMBYN_PNRSN_CHNLLNG_MT_NG_SNG"].fillna(0)
    df_sya["SLD_TSTNDNG_PRNCPLS_PNYLRN_PMBYN_BRSM_TTL"] = df_sya["KGTN_PMBYN_PNRSN_CHNLLNG_TTL"].fillna(0)

    # Nominal Derivatif Lainnya
    df_sya["LNNY_TTL"] = df_sya["NMNL_SPT_TTL"].fillna(0)

    return df_sya

## Rekapitulasi Tenaga Kerja Berdasarkan Tingkat Divisi
def transform_tingkat_divisi(df_sya):
    # Ganti _TNG_KRJ_TNG_KRJ_ → _TNG_KRJ_
    df_sya.rename(
        columns={col: col.replace('_TNG_KRJ_TNG_KRJ_', '_TNG_KRJ_') 
                 for col in df_sya.columns if '_TNG_KRJ_TNG_KRJ_' in col},
        inplace=True
    )
    # Isi NaN dengan 0 hanya untuk kolom _TNG_KRJ_
    cols_to_fill = [col for col in df_sya.columns if '_TNG_KRJ_' in col]
    df_sya[cols_to_fill] = df_sya[cols_to_fill].fillna(0)

    return df_sya

## Rekapitulasi Tenaga Kerja Berdasarkan Tingkat Pendidikan
def transform_tingkat_pendidikan(df_sya):
    # Ganti akhiran _LK_LK → _LKLK
    df_sya.rename(
        columns={col: col.replace('_LK_LK', '_LKLK') for col in df_sya.columns if col.endswith('_LK_LK')},
        inplace=True
    )
    # Ganti __TTL → _TTL_ (untuk kolom total)
    df_sya.rename(
        columns={col: col.replace('__TTL', '_TTL_') for col in df_sya.columns if '__TTL' in col},
        inplace=True
    )
    # Isi NaN dengan 0 untuk semua kolom yang diganti
    cols_to_fill = [col for col in df_sya.columns if col.endswith('_LKLK') or '_TTL_' in col]
    df_sya[cols_to_fill] = df_sya[cols_to_fill].fillna(0)

    return df_sya

## Laporan Kesesuaian Aset dan Liabilitas
def transform_lkal(df_sya):
    """
    Transformasi seragam Arus Kas Syariah dan LKAL menggunakan mapping dictionary.
    Semua penggantian kolom dilakukan satu per satu (tanpa penjumlahan).
    """
    
    # Mapping lengkap dan seragam untuk semua transformasi
    mapping_seragam = {
        # ----------------------------- ARUS KAS MASUK -----------------------------
        # Investasi Syariah
        "RS_KS_MSK_DR_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH_NDNSN_RPH": 
        "RS_KS_MSK_DR_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_NDNSN_RPH",
        
        "RS_KS_MSK_DR_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH_MT_NG_SNG": 
        "RS_KS_MSK_DR_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_MT_NG_SNG",
        
        "RS_KS_MSK_DR_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH_TTL": 
        "RS_KS_MSK_DR_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_TTL",
        
        # Jasa Berbasis Fee
        "RS_KS_MSK_DR_KGTN_JS_BRBSS_F_NDNSN_RPH": 
        "RS_KS_MSK_DR_KGTN_JS_BRBSS_MBL_HSL_NDNSN_RPH",
        
        "RS_KS_MSK_DR_KGTN_JS_BRBSS_F_MT_NG_SNG": 
        "RS_KS_MSK_DR_KGTN_JS_BRBSS_MBL_HSL_MT_NG_SNG",
        
        "RS_KS_MSK_DR_KGTN_JS_BRBSS_F_TTL": 
        "RS_KS_MSK_DR_KGTN_JS_BRBSS_MBL_HSL_TTL",
        
        # ----------------------------- ARUS KAS KELUAR INVESTASI -----------------------------
        "RS_KS_KLR_NTK_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH__NDNSN_RPH": 
        "RS_KS_KLR_NTK_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_NDNSN_RPH",
        
        "RS_KS_KLR_NTK_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH__MT_NG_SNG": 
        "RS_KS_KLR_NTK_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_MT_NG_SNG",
        
        "RS_KS_KLR_NTK_KGTN_NVSTS_BRDSRKN_PRNSP_SYRH__TTL": 
        "RS_KS_KLR_NTK_KGTN_PNYRTN_SHM_BRDSRKN_PRNSP_SYRH_TTL",
        
        # ----------------------------- ARUS KAS KELUAR JASA -----------------------------
        "RS_KS_KLR_NTK_KGTN_JS_BRBSS_F_NDNSN_RPH": 
        "RS_KS_KLR_DR_KGTN_JS_BRBSS_MBL_HSL_NDNSN_RPH",
        
        "RS_KS_KLR_NTK_KGTN_JS_BRBSS_F_MT_NG_SNG": 
        "RS_KS_KLR_DR_KGTN_JS_BRBSS_MBL_HSL_MT_NG_SNG",
        
        "RS_KS_KLR_NTK_KGTN_JS_BRBSS_F_TTL": 
        "RS_KS_KLR_DR_KGTN_JS_BRBSS_MBL_HSL_TTL",
        
        # ----------------------------- ARUS KAS KELUAR PENDAANAN -----------------------------
        "RS_KS_KLR_NTK_PNDNN_BRDSRKN_PRNSP_SYRH_NDNSN_RPH": 
        "RS_KS_KLR_NTK_PMBYRN_PKK_PNJMN_DN_SRT_BRHRG_YNG_DTRBTKN_NDNSN_RPH",
        
        "RS_KS_KLR_NTK_PNDNN_BRDSRKN_PRNSP_SYRH_MT_NG_SNG": 
        "RS_KS_KLR_NTK_PMBYRN_PKK_PNJMN_DN_SRT_BRHRG_YNG_DTRBTKN_MT_NG_SNG",
        
        "RS_KS_KLR_NTK_PNDNN_BRDSRKN_PRNSP_SYRH_TTL": 
        "RS_KS_KLR_NTK_PMBYRN_PKK_PNJMN_DN_SRT_BRHRG_YNG_DTRBTKN_TTL",
        
        # ----------------------------- LKAL MAPPING -----------------------------
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PD_STRTP_DNT_PNGMBNGN_SH_NT_VLS_LBH_DR_3_BLN_DN_KRNG_DR_T_SM_DNGN_6_BLN":
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTP_NT_MT_NG_SNG_LBH_DR_3_BLN_DN_KRNG_DR_T_SM_DNGN_6_BLN",
        
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PD_STRTP_DNT_PNGMBNGN_SH_NT_VLS_LBH_DR_6_BLN_DN_KRNG_DR_T_SM_DNGN_1_THN":
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTP_NT_MT_NG_SNG_LBH_DR_6_BLN_DN_KRNG_DR_T_SM_DNGN_1_THN",
        
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PD_STRTP_DNT_PNGMBNGN_SH_NT_VLS_LBH_DR_1_THN_DN_KRNG_DR_T_SM_DNGN_5_THN":
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTP_NT_MT_NG_SNG_LBH_DR_1_THN_DN_KRNG_DR_T_SM_DNGN_5_THN",
        
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PD_STRTP_DNT_PNGMBNGN_SH_NT_VLS_LBH_DR_5_THN_DN_KRNG_DR_T_SM_DNGN_10_THN":
        "PMBLN_SKK_T_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTP_NT_MT_NG_SNG_LBH_DR_5_THN_DN_KRNG_DR_T_SM_DNGN_10_THN",
        
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_SKK_BLGS_SYRH_YNG_DTRBTKN_STRTP_PNGMBNGN_SH_VLS_LBH_DR_3_BLN_DN_KRNG_DR_T_SM_DNGN_6_BLN":
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTPSH_VLS_LBH_DR_3_BLN_DN_KRNG_DR_T_SM_DNGN_6_BLN",
        
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_SKK_BLGS_SYRH_YNG_DTRBTKN_STRTP_PNGMBNGN_SH_VLS_LBH_DR_6_BLN_DN_KRNG_DR_T_SM_DNGN_1_THN":
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTPSH_VLS_LBH_DR_6_BLN_DN_KRNG_DR_T_SM_DNGN_1_THN",
        
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_SKK_BLGS_SYRH_YNG_DTRBTKN_STRTP_PNGMBNGN_SH_VLS_LBH_DR_1_THN_DN_KRNG_DR_T_SM_DNGN_5_THN":
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTPSH_VLS_LBH_DR_1_THN_DN_KRNG_DR_T_SM_DNGN_5_THN",
        
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_SKK_BLGS_SYRH_YNG_DTRBTKN_STRTP_PNGMBNGN_SH_VLS_LBH_DR_5_THN_DN_KRNG_DR_T_SM_DNGN_10_THN":
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTPSH_VLS_LBH_DR_5_THN_DN_KRNG_DR_T_SM_DNGN_10_THN",
        
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_SKK_BLGS_SYRH_YNG_DTRBTKN_STRTP_PNGMBNGN_SH_VLS_KRNG_DR_T_SM_DNGN_3_BLN":
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTPSH_VLS_KRNG_DR_T_SM_DNGN_3_BLN",
        
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_SKK_BLGS_SYRH_YNG_DTRBTKN_STRTP_PNGMBNGN_SH_VLS_LBH_DR_10_THN":
        "CDNGN_PNYSHN_PNGHPSN_ST_PRDKTF_PMBLN_BLGS_SYRH_YNG_DTRBTKN_PSNGN_SH_PD_STRTPSH_VLS_LBH_DR_10_THN",
    }
    
    # Terapkan mapping secara seragam
    for kolom_target, kolom_sumber in mapping_seragam.items():
        df_sya[kolom_target] = df_sya[kolom_sumber].fillna(0)
    
    return df_sya

# Config Laporan
def get_app_dir():
    # Jika dijalankan sebagai .exe
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)

    # Jika dijalankan sebagai script .py
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        # Jika dijalankan di Jupyter / IPython
        return os.getcwd()

APP_DIR = get_app_dir()
TEMPLATE_DIR = os.path.join(APP_DIR, "templates")
DATA_REF_DIR = os.path.join(APP_DIR, "data") 

LAPORAN_CONFIG = {
    "neraca": {
        "form": "1100",
        "template": os.path.join(TEMPLATE_DIR, "Template_Neraca_PMV.xlsx"),
        "sheet": "Neraca Gabungan",
        "transform": None,
        "output_folder" : "03. Neraca"
    },
    "laba_rugi": {
        "form": "1200",
        "template": os.path.join(TEMPLATE_DIR, "Template_Laba_Rugi_PMV.xlsx"),
        "sheet": "Laba Rugi Gabungan",
        "transform": transform_laba_rugi,
        "output_folder" : "05. Laba Rugi"
    },
    "arus_kas": {
        "form": "1300",
        "template": os.path.join(TEMPLATE_DIR, "Template_Arus_Kas_PMV.xlsx"),
        "sheet": "Arus Kas Gabungan",
        "transform": transform_arus_kas,
        "output_folder" : "06. Arus Kas"
    },
    "rekening_administratif": {
        "form": "1110",
        "template": os.path.join(TEMPLATE_DIR, "Template_Rekening_Administratif_PMV.xlsx"),
        "sheet": "Rekening Administratif Gabungan",
        "transform": transform_rekening_administratif,
        "output_folder" : "07. Rekening Administratif"
    },
     "laporan_kesesuaian_aset_dan_liabilitas": {
        "form": "5310",
        "template": os.path.join(TEMPLATE_DIR, "Template_LKAL_PMV.xlsx"),
        "sheet": "LKAL Gabungan",
        "transform": transform_lkal,
        "output_folder" : "08. Laporan Kesesuaian Aset dan Liabilitas"
    },
     "tingkat_divisi": {
        "form": "0043",
        "template": os.path.join(TEMPLATE_DIR, "Template_Tingkat_Divisi_PMV.xlsx"),
        "sheet": "TK Berdasarkan Tingkat Divisi",
        "transform": transform_tingkat_divisi,
        "output_folder" : "10. TK Berdasarkan Tingkat Divisi"
    },
     "tingkat_pendidikan": {
        "form": "0041",
        "template": os.path.join(TEMPLATE_DIR, "Template_Tingkat_Pendidikan_PMV.xlsx"),
        "sheet": "TKBerdasarkanTingkatPendidikan",
        "transform": transform_tingkat_pendidikan,
        "output_folder" : "09. TK Berdasarkan Tingkat Pendidikan"
    }
}

# Fungsi Laporan
## NPF
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter


def laporan_npf(template_path, output_path, df_perusahaan, df_piutang):
    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 3
    NUM_FORMAT     = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    PERCENT_FORMAT = '0.00%'

    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL   = PatternFill("solid", fgColor="FF99FF")

    thin = Side(style="thin")
    BORDER_THIN = Border(left=thin, right=thin, top=thin, bottom=thin)

    FORM_SHEET_MAP = {
        2110: "NPF PS",
        2120: "NPF OK", 
        2130: "NPF SUS",
        2140: "NPF PUP"
    }

    # ✅ NEW: NPF Perusahaan Sheet
    NPF_PERUSAHAAN = "NPF Perusahaan"

    PIUTANG_COL = {
        "SF:e2": 6,   # F
        "SF:e12": 7,  # G
        "SF:e9": 8,   # H
        "SF:e3": 9,   # I
        "SF:e4": 10   # J
    }

    CKPN_COL = {
        "SF:e2": 13,  # M
        "SF:e12": 14, # N
        "SF:e9": 15,  # O
        "SF:e3": 16,  # P
        "SF:e4": 17   # Q
    }

    # =====================================================
    # UTILITAS
    # =====================================================
    def apply_border(ws, last_row, end_col):
        for r in range(2, last_row + 1):
            for c in range(4, end_col + 1):
                ws.cell(r, c).border = BORDER_THIN

    def apply_number(ws, last_row, end_col, is_ps=True, is_perusahaan=False):
        """Format NUM_FORMAT untuk data + formula columns"""
        for r in range(START_ROW, last_row + 1):
            # Data columns (F-J, M-Q, dll)
            start_col = 6
            if is_perusahaan:
                end_data_col = 21  # Sampai U untuk Perusahaan
            else:
                end_data_col = end_col
            
            for c in range(start_col, end_data_col + 1):
                cell = ws.cell(r, c)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = NUM_FORMAT
            
            # Formula columns
            if is_ps:
                # PS: K dan L
                ws[f"K{r}"].number_format = NUM_FORMAT
                ws[f"L{r}"].number_format = NUM_FORMAT
            elif is_perusahaan:
                # ✅ FIXED: Perusahaan: K,L,R,S (NUM) + T,U (PERCENT)
                ws[f"K{r}"].number_format = NUM_FORMAT
                ws[f"L{r}"].number_format = NUM_FORMAT
                ws[f"R{r}"].number_format = NUM_FORMAT
                ws[f"S{r}"].number_format = NUM_FORMAT
                # T,U sudah di-set PERCENT_FORMAT di fill_npf_perusahaan
            else:
                # Non-PS: K,L,R,S
                ws[f"K{r}"].number_format = NUM_FORMAT
                ws[f"L{r}"].number_format = NUM_FORMAT
                ws[f"R{r}"].number_format = NUM_FORMAT
                ws[f"S{r}"].number_format = NUM_FORMAT

    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]
            if p.get("JNS_PRSHN") == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH
            row += 1
        return row - 1

    def apply_total(ws, last_row, is_ps):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)

        end_col = 13 if is_ps else 21  # ✅ Sesuai kode Anda

        for c in range(1, end_col + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL

        # PIUTANG (F-J)
        for c in range(6, 11):
            col = get_column_letter(c)
            ws.cell(total_row, c).value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            ws.cell(total_row, c).number_format = NUM_FORMAT

        ws[f"K{total_row}"] = f"=SUM(H{total_row}:J{total_row})"
        ws[f"K{total_row}"].number_format = NUM_FORMAT
        ws[f"L{total_row}"] = f"=SUM(F{total_row}:J{total_row})"
        ws[f"L{total_row}"].number_format = NUM_FORMAT

        if is_ps:
            ws[f"M{total_row}"] = f"=IFERROR(K{total_row}/L{total_row},0)"
            ws[f"M{total_row}"].number_format = PERCENT_FORMAT
        else:
            for c in range(13, 18):
                col = get_column_letter(c)
                ws.cell(total_row, c).value = f"=SUM({col}{START_ROW}:{col}{last_row})"
                ws.cell(total_row, c).number_format = NUM_FORMAT

            ws[f"R{total_row}"] = f"=SUM(O{total_row}:Q{total_row})"
            ws[f"R{total_row}"].number_format = NUM_FORMAT
            ws[f"S{total_row}"] = f"=SUM(M{total_row}:Q{total_row})"
            ws[f"S{total_row}"].number_format = NUM_FORMAT
            ws[f"T{total_row}"] = f"=IFERROR(K{total_row}/L{total_row},0)"
            ws[f"T{total_row}"].number_format = PERCENT_FORMAT
            ws[f"U{total_row}"] = f"=IFERROR(IF((K{total_row}-R{total_row})/L{total_row}<0,0,(K{total_row}-R{total_row})/L{total_row}),0)"
            ws[f"U{total_row}"].number_format = PERCENT_FORMAT

        return total_row, end_col

    # ✅ FIXED: NPF Perusahaan Formulas - T,U sudah PERCENT_FORMAT
    def fill_npf_perusahaan(ws, last_row):
        """F3 = 'NPF PS'!F3 + 'NPF OK'!F3 + ... dst"""
        sheets = ['NPF PS', 'NPF OK', 'NPF SUS', 'NPF PUP']
        
        for r in range(START_ROW, last_row + 1):
            # PIUTANG F-J: sum dari 4 sheets
            for c in range(6, 11):  # F-J
                col_letter = get_column_letter(c)
                formula = "='{}'!{}{}+'{}'!{}{}+'{}'!{}{}+'{}'!{}{}".format(
                    sheets[0], col_letter, r, sheets[1], col_letter, r, 
                    sheets[2], col_letter, r, sheets[3], col_letter, r
                )
                ws[f"{col_letter}{r}"].value = formula
                ws[f"{col_letter}{r}"].number_format = NUM_FORMAT
            
            # CKPN M-Q: sum dari 3 sheets (tanpa PS)
            for c in range(13, 18):  # M-Q
                col_letter = get_column_letter(c)
                formula = "='{}'!{}{}+'{}'!{}{}+'{}'!{}{}".format(
                    sheets[1], col_letter, r, sheets[2], col_letter, r, sheets[3], col_letter, r
                )
                ws[f"{col_letter}{r}"].value = formula
                ws[f"{col_letter}{r}"].number_format = NUM_FORMAT
            
            # Formula totals
            ws[f"K{r}"] = f"=SUM(H{r}:J{r})"
            ws[f"L{r}"] = f"=SUM(F{r}:J{r})"
            ws[f"R{r}"] = f"=SUM(O{r}:Q{r})"
            ws[f"S{r}"] = f"=SUM(M{r}:Q{r})"
            ws[f"T{r}"] = f"=IFERROR(K{r}/L{r},0)"
            ws[f"T{r}"].number_format = PERCENT_FORMAT  # ✅ PERCENT
            ws[f"U{r}"] = f"=IFERROR(IF((K{r}-R{r})/L{r}<0,0,(K{r}-R{r})/L{r}),0)"
            ws[f"U{r}"].number_format = PERCENT_FORMAT  # ✅ PERCENT

    # =====================================================
    # REKAP DATA
    # =====================================================
    rekap_piu  = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    rekap_ckpn = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for _, r in df_piutang.iterrows():
        form  = r.get("FORM")
        ident = r.get("IDENTIFIER")
        klts  = r.get("KLTS")

        if form in FORM_SHEET_MAP and pd.notna(ident) and klts in PIUTANG_COL:
            rekap_piu[form][ident][klts] += r.get("PIUTANG", 0)
            if form != 2110:
                rekap_ckpn[form][ident][klts] += r.get("CKPN", 0)

    # =====================================================
    # PROSES PER SHEET
    # =====================================================
    wb = load_workbook(template_path, data_only=False)

    # 1. Proses NPF PS/OK/SUS/PUP
    for form, sheet_name in FORM_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        last_row = fill_base(ws)
        is_ps = (form == 2110)

        # Isi data per baris
        for r in range(START_ROW, last_row + 1):
            ident = ws.cell(r, 2).value
            if not ident:
                continue

            for k, c in PIUTANG_COL.items():
                ws.cell(r, c).value = rekap_piu.get(form, {}).get(ident, {}).get(k, 0)

            ws[f"K{r}"] = f"=SUM(H{r}:J{r})"
            ws[f"L{r}"] = f"=SUM(F{r}:J{r})"

            if is_ps:
                ws[f"M{r}"] = f"=IFERROR(K{r}/L{r},0)"
                ws[f"M{r}"].number_format = PERCENT_FORMAT
            else:
                for k, c in CKPN_COL.items():
                    ws.cell(r, c).value = rekap_ckpn.get(form, {}).get(ident, {}).get(k, 0)
                
                ws[f"R{r}"] = f"=SUM(O{r}:Q{r})"
                ws[f"S{r}"] = f"=SUM(M{r}:Q{r})"
                ws[f"T{r}"] = f"=IFERROR(K{r}/L{r},0)"
                ws[f"T{r}"].number_format = PERCENT_FORMAT
                ws[f"U{r}"] = f"=IFERROR(IF((K{r}-R{r})/L{r}<0,0,(K{r}-R{r})/L{r}),0)"
                ws[f"U{r}"].number_format = PERCENT_FORMAT

        total_row, end_col = apply_total(ws, last_row, is_ps)
        apply_number(ws, last_row, end_col, is_ps)
        apply_border(ws, last_row, end_col)

    # 2. Proses NPF Perusahaan
    if NPF_PERUSAHAAN not in wb.sheetnames:
        print(f"⚠️  Sheet '{NPF_PERUSAHAAN}' tidak ditemukan di template")
    else:
        ws_perusahaan = wb[NPF_PERUSAHAAN]
        last_row = fill_base(ws_perusahaan)
        
        # ✅ FORMULA SUM antar sheets
        fill_npf_perusahaan(ws_perusahaan, last_row)
        
        # Total row
        total_row, end_col = apply_total(ws_perusahaan, last_row, is_ps=False)
        apply_number(ws_perusahaan, last_row, end_col, is_ps=False, is_perusahaan=True)
        apply_border(ws_perusahaan, last_row, end_col)

    # =====================================================
    # SAVE DENGAN FORMULA AKTIF
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()
    
    wb_final = load_workbook(output_path, data_only=False)
    for ws in wb_final.worksheets:
        if hasattr(ws.sheet_properties, 'calculationMode'):
            ws.sheet_properties.calculationMode = "automatic"
    wb_final.save(output_path)
    wb_final.close()
## Sektor Ekonomi
### 1. Sektor Ekonomi vs. NPF
def laporan_sektor_vs_npf(template_path, output_path, df_piutang):

    wb = load_workbook(template_path)

    MAP_KLTS_PIUTANG = {
        "SF:e2":  "C",
        "SF:e12": "D",
        "SF:e9":  "E",
        "SF:e3":  "F",
        "SF:e4":  "G"
    }

    MAP_KLTS_CKPN = {
        "SF:e2":  "J",
        "SF:e12": "K",
        "SF:e9":  "L",
        "SF:e3":  "M",
        "SF:e4":  "N"
    }

    FORM_TO_SHEET = {
    "2110": "PS",
    "2120": "OK",
    "2130": "SUS",
    "2140": "PUP"
    }   

    rekap_piutang = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    rekap_ckpn    = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for _, r in df_piutang.iterrows():
        form   = str(r["FORM"])
        sektor = r["SKTR_KNM_LPNGN_SH"]
        klts   = r["KLTS"]

        if pd.notna(r.get("PIUTANG")):
            rekap_piutang[form][sektor][klts] += r["PIUTANG"]

        if pd.notna(r.get("CKPN")):
            rekap_ckpn[form][sektor][klts] += r["CKPN"]

    for form, sheet in FORM_TO_SHEET.items():
        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]

        for row in range(3, ws.max_row + 1):
            sektor = ws[f"B{row}"].value
            if not sektor:
                continue

            data_piu = rekap_piutang.get(form, {}).get(sektor)
            data_ckp = rekap_ckpn.get(form, {}).get(sektor)

            if not data_piu:
                continue

            for klts, col in MAP_KLTS_PIUTANG.items():
                if klts in data_piu:
                    ws[f"{col}{row}"] = data_piu[klts]

            if sheet != "PS" and data_ckp:
                for klts, col in MAP_KLTS_CKPN.items():
                    if klts in data_ckp:
                        ws[f"{col}{row}"] = data_ckp[klts]

    wb.save(output_path)

### 2. Sektor Ekonomi vs. Jenis Usaha
def laporan_sektor_vs_jenis_usaha(template_path, output_path, df_piutang):

    wb = load_workbook(template_path)
    ws = wb["FIX"]

    MAP_AKUN_COL = {
        "AK:e20": "D",
        "AK:e21": "E",
        "AK:e22": "F",
        "AK:e23": "G",
        "AK:e24": "I",
        "AK:e25": "J",
        "AK:e26": "K",
        "AK:e27": "L"
    }

    rekap = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        sektor = r["SKTR_KNM_LPNGN_SH"]
        akun   = r["JNS_KGTN_SH_PRSHN_MDL_VNTR"]

        if akun in MAP_AKUN_COL:
            rekap[sektor][akun] += r["PIUTANG"]

    for row in range(1, ws.max_row + 1):
        sektor = ws[f"B{row}"].value
        if sektor not in rekap:
            continue

        for akun, col in MAP_AKUN_COL.items():
            if akun in rekap[sektor]:
                ws[f"{col}{row}"] = rekap[sektor][akun]

    wb.save(output_path)

### 3. Sektor Ekonomi vs. Identifier
def laporan_sektor_vs_identifier(template_path, output_path, df_perusahaan, df_piutang):
    """
    Mengisi laporan lokasi vs identifier untuk:
    - 34 Prov
    - 38 Prov
    (logika sama, jumlah perusahaan dinamis, TOTAL otomatis)
    Template mulai dari 1 kolom C, kolom baru otomatis copy formula+style
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment
    from copy import copy
    from collections import defaultdict
    from openpyxl.formula.translate import Translator
    import os

    # =====================================================
    # KONFIGURASI
    # =====================================================
    SHEET_NAMES = ["FIX"]
    START_COL = 3        # Kolom C
    HEADER_ROW_1 = 1
    HEADER_ROW_2 = 2
    START_ROW_DATA = 3

    # Helper function untuk copy kolom lengkap (formula + style)
    def copy_col_with_style_and_formula(ws, src_col_idx, dst_col_idx, row_start=1, row_end=None):
        if row_end is None:
            row_end = ws.max_row
        
        for r in range(row_start, row_end + 1):
            src = ws.cell(row=r, column=src_col_idx)
            dst = ws.cell(row=r, column=dst_col_idx)

            # Copy style lengkap
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)

            # Copy value (kalau formula, translate referensi)
            v = src.value
            if isinstance(v, str) and v.startswith("="):
                # Formula otomatis adjust: C3=SUM(C4:C5) → D3=SUM(D4:D5)
                dst.value = Translator(v, origin=src.coordinate).translate_formula(dst.coordinate)
            else:
                dst.value = v

    # Load workbook
    wb = load_workbook(template_path)
    ident_list = df_perusahaan["IDENTIFIER"].tolist()

    # =====================================================
    # REKAP DATA (LOKASI × IDENTIFIER)
    # =====================================================
    rekap = defaultdict(lambda: defaultdict(float))
    for _, r in df_piutang.iterrows():
        rekap[r["SKTR_KNM_LPNGN_SH"]][r["IDENTIFIER"]] += r["PIUTANG"]

    # =====================================================
    # PROSES SETIAP SHEET
    # =====================================================
    for sheet_name in SHEET_NAMES:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        # -------------------------------------------------
        # 1️⃣ TAMBAH KOLOM PERUSAHAAN (tanpa hapus kolom)
        # -------------------------------------------------
        required_cols = len(ident_list)
        current_company_cols = ws.max_column - START_COL + 1  # Kolom perusahaan saat ini

        # Tambah kolom sampai cukup
        while current_company_cols < required_cols:
            insert_at = START_COL + current_company_cols  # Posisi kolom baru
            ws.insert_cols(insert_at)
            
            # Copy dari kolom sebelumnya (template kolom C akan propagate)
            src_col_idx = insert_at - 1
            dst_col_idx = insert_at
            copy_col_with_style_and_formula(ws, src_col_idx, dst_col_idx)
            
            current_company_cols += 1

        # -------------------------------------------------
        # 2️⃣ ISI HEADER IDENTIFIER & NAMA PERUSAHAAN
        # -------------------------------------------------
        for i, p in enumerate(df_perusahaan.itertuples(index=False)):
            col = get_column_letter(START_COL + i)
            
            # Header 1: IDENTIFIER → CENTER
            header1_cell = ws[f"{col}{HEADER_ROW_1}"]
            header1_cell.value = p.IDENTIFIER
            header1_cell.alignment = Alignment(horizontal='center')
            
            # Header 2: NAMA PERUSAHAAN → tetap original
            ws[f"{col}{HEADER_ROW_2}"] = p.NM_SBTN_PRSHN

        # -------------------------------------------------
        # 3️⃣ ISI DATA (hanya override VALUE data cells, formula aman)
        # -------------------------------------------------
        for row in range(START_ROW_DATA, ws.max_row + 1):
            lokasi = ws[f"B{row}"].value
            if not lokasi:
                continue

            for i, ident in enumerate(ident_list):
                col = get_column_letter(START_COL + i)
                cell = ws[f"{col}{row}"]
                
                # **CEK FORMULA DULU**
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue  # Formula → Skip
                
                # Data → Isi (0 jadi kosong/None)
                data_value = rekap[lokasi].get(ident, 0) if lokasi in rekap else 0
                cell.value = data_value

        # -------------------------------------------------
        # 4️⃣ BUAT KOLOM TOTAL
        # -------------------------------------------------
        total_col_idx = START_COL + len(ident_list)
        total_col = get_column_letter(total_col_idx)
        ws.insert_cols(total_col_idx)

        # Copy style dari kolom kiri terakhir
        last_data_col_idx = total_col_idx - 1
        for r in range(1, ws.max_row + 1):
            src = ws.cell(r, last_data_col_idx)
            dst = ws.cell(r, total_col_idx)
            
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format

        ws[f"{total_col}{HEADER_ROW_2}"] = "TOTAL"
        
        # Formula TOTAL per baris
        last_data_col_letter = get_column_letter(total_col_idx - 1)
        for r in range(START_ROW_DATA, ws.max_row + 1):
            ws[f"{total_col}{r}"] = f"=SUM(C{r}:{last_data_col_letter}{r})"

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()

### 4. Identifier vs. Sektor Ekonomi
def laporan_identifier_vs_sektor(template_path, output_path, df_perusahaan, df_piutang):
    from collections import defaultdict
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Border, Side
    from openpyxl.utils import get_column_letter
    import os
    import pandas as pd

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 3
    START_COL = 4      # D
    DATA_END  = 26     # Z
    END_COL   = 27     # AA

    NUM_FORMAT     = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    PERCENT_FORMAT = '0.00%'

    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL   = PatternFill("solid", fgColor="FF99FF")

    thin = Side(style="thin")
    BORDER_THIN = Border(left=thin, right=thin, top=thin, bottom=thin)

    # =====================================================
    # UTILITAS
    # =====================================================
    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]

            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH

            ws[f"AA{row}"] = f"=SUM(D{row}:Z{row})"
            row += 1
        return row - 1

    def apply_border(ws, last_row):
        for r in range(START_ROW, last_row + 1):
            for c in range(START_COL, END_COL + 1):
                ws.cell(r, c).border = BORDER_THIN

    def apply_number(ws, last_row, fmt):
        for r in range(START_ROW, last_row + 1):
            for c in range(START_COL, END_COL + 1):
                cell = ws.cell(r, c)
                if cell.value is None:
                    cell.value = 0
                cell.number_format = fmt

    def apply_total_row(ws, last_row, fmt):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)

        for c in range(1, END_COL + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL

        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = fmt

        return total_row

    def fill_value_sheet(ws, df_src):
        sektor_col = {
            ws.cell(1, c).value: c
            for c in range(START_COL, END_COL + 1)
            if str(ws.cell(1, c).value).startswith("SE:")
        }

        rekap = defaultdict(lambda: defaultdict(float))
        for _, r in df_src.iterrows():
            if pd.notna(r["SEKTOR_KAT"]):
                rekap[r["IDENTIFIER"]][r["SEKTOR_KAT"]] += r["PIUTANG"]

        last_row = fill_base(ws)

        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value
            for sektor, val in rekap.get(ident, {}).items():
                if sektor in sektor_col:
                    ws.cell(r, sektor_col[sektor]).value = val

        apply_number(ws, last_row, NUM_FORMAT)
        apply_border(ws, last_row)
        apply_total_row(ws, last_row, NUM_FORMAT)

        return last_row

    # =====================================================
    # LOAD TEMPLATE
    # =====================================================
    wb = load_workbook(template_path)

    # =====================================================
    # TOTAL
    # =====================================================
    fill_value_sheet(wb["Total"], df_piutang)

    # =====================================================
    # NON PERFORM
    # =====================================================
    df_np = df_piutang[df_piutang["KLTS"].isin(["SF:e9", "SF:e3", "SF:e4"])]
    fill_value_sheet(wb["NonPerform"], df_np)

    # =====================================================
    # NPF (%)
    # =====================================================
    ws_npf = wb["NPF"]
    last_row = fill_base(ws_npf)

    for r in range(START_ROW, last_row + 1):
        for c in range(START_COL, DATA_END + 1):
            col = get_column_letter(c)
            ws_npf.cell(r, c).value = (
                f"=IFERROR(NonPerform!{col}{r}/Total!{col}{r},0)"
            )

        ws_npf[f"AA{r}"] = f"=IFERROR(NonPerform!AA{r}/Total!AA{r},0)"

    total_row = last_row + 2
    ws_npf[f"C{total_row}"] = "Total"
    ws_npf[f"C{total_row}"].font = Font(bold=True)

    for c in range(1, END_COL + 1):
        ws_npf.cell(total_row, c).fill = FILL_TOTAL

    for c in range(START_COL, DATA_END + 1):
        col = get_column_letter(c)
        ws_npf.cell(total_row, c).value = (
            f"=IFERROR(NonPerform!{col}{total_row}/Total!{col}{total_row},0)"
        )
        ws_npf.cell(total_row, c).font = Font(bold=True)

    ws_npf[f"AA{total_row}"] = (
        f"=IFERROR(NonPerform!AA{total_row}/Total!AA{total_row},0)"
    )
    ws_npf[f"AA{total_row}"].font = Font(bold=True)

    for r in list(range(START_ROW, last_row + 1)) + [total_row]:
        for c in range(START_COL, END_COL + 1):
            ws_npf.cell(r, c).number_format = PERCENT_FORMAT

    apply_border(ws_npf, last_row)

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
## UMKM
### 1. UMKM vs. NPF
def laporan_umkm_vs_npf(template_path, output_path, df_piutang):
    START_ROW = 2

    FORM_SHEET = {
        "2110": "PS",
        "2120": "OK",
        "2130": "SUS",
        "2140": "PUP"
    }

    PIUTANG_COL = {
        "SF:e2": 3,
        "SF:e12": 4,
        "SF:e9": 5,
        "SF:e3": 6,
        "SF:e4": 7
    }

    CKPN_COL = {
        "SF:e2": 10,
        "SF:e12": 11,
        "SF:e9": 12,
        "SF:e3": 13,
        "SF:e4": 14
    }

    wb = load_workbook(template_path)

    for form, sheet in FORM_SHEET.items():

        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]
        df_form = df_piutang[df_piutang["FORM"].astype(str).str.startswith(form)]

        rekap_piu  = defaultdict(lambda: defaultdict(float))
        rekap_ckpn = defaultdict(lambda: defaultdict(float))

        for _, r in df_form.iterrows():
            ktgr = r.get("KTGR_SH_DBTR")
            klts = r.get("KLTS")

            if pd.notna(ktgr) and klts in PIUTANG_COL:
                rekap_piu[ktgr][klts]  += r.get("PIUTANG", 0)
                rekap_ckpn[ktgr][klts] += r.get("CKPN", 0)

        row = START_ROW
        while ws.cell(row, 2).value is not None:

            sandi = ws.cell(row, 2).value
            ws.cell(row, 1).value = sandi
            ws.cell(row, 2).value = sandi

            for k, c in PIUTANG_COL.items():
                ws.cell(row, c).value = rekap_piu.get(sandi, {}).get(k, 0)

            for k, c in CKPN_COL.items():
                ws.cell(row, c).value = rekap_ckpn.get(sandi, {}).get(k, 0)

            row += 1

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 2. UMKM vs. Jenis Usaha
def laporan_umkm_vs_jenis_usaha(template_path, output_path, df_piutang):
    from collections import defaultdict
    from openpyxl import load_workbook
    import os
    import pandas as pd

    START_ROW = 3

    # =========================
    # MAP JENIS USAHA → KOLOM
    # =========================
    USAHA_COL = {
        "AK:e20": 4,   # D
        "AK:e21": 5,   # E
        "AK:e22": 6,   # F
        "AK:e23": 7,   # G
        "AK:e24": 9,   # I
        "AK:e25": 10,  # J
        "AK:e26": 11,  # K
        "AK:e27": 12   # L
    }

    wb = load_workbook(template_path)
    ws = wb["Fix"]   # sesuai template

    # =========================
    # REKAP DATA
    # =========================
    rekap = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        ktgr  = r.get("KTGR_SH_DBTR")
        usaha = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
        nilai = r.get("PIUTANG", 0)

        if pd.notna(ktgr) and usaha in USAHA_COL:
            rekap[ktgr][usaha] += nilai

    # =========================
    # ISI KE TEMPLATE
    # =========================
    row = START_ROW
    while ws.cell(row, 2).value is not None:

        sandi = ws.cell(row, 2).value
        ws.cell(row, 1).value = sandi
        ws.cell(row, 2).value = sandi

        for usaha, col in USAHA_COL.items():
            ws.cell(row, col).value = rekap.get(sandi, {}).get(usaha, 0)

        row += 1

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 3. UMKM vs. Jenis Usaha (Kontrak)
def laporan_umkm_vs_jenis_usaha_kontrak(template_path, output_path, df_piutang):
    from collections import defaultdict
    from openpyxl import load_workbook
    import os
    import pandas as pd

    START_ROW = 3

    # =========================
    # MAP JENIS USAHA → KOLOM
    # =========================
    USAHA_COL = {
        "AK:e20": 4,   # D
        "AK:e21": 5,   # E
        "AK:e22": 6,   # F
        "AK:e23": 7,   # G
        "AK:e24": 9,   # I
        "AK:e25": 10,  # J
        "AK:e26": 11,  # K
        "AK:e27": 12   # L
    }

    wb = load_workbook(template_path)
    ws = wb["Fix"]   # sesuai template

    # =========================
    # REKAP DATA
    # =========================
    rekap = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        ktgr  = r.get("KTGR_SH_DBTR")
        usaha = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
        nilai = r.get("KONTRAK", 0)

        if pd.notna(ktgr) and usaha in USAHA_COL:
            rekap[ktgr][usaha] += nilai

    # =========================
    # ISI KE TEMPLATE
    # =========================
    row = START_ROW
    while ws.cell(row, 2).value is not None:

        sandi = ws.cell(row, 2).value
        ws.cell(row, 1).value = sandi
        ws.cell(row, 2).value = sandi

        for usaha, col in USAHA_COL.items():
            ws.cell(row, col).value = rekap.get(sandi, {}).get(usaha, 0)

        row += 1

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 4. Identifier vs. UMKM
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def laporan_identifier_vs_umkm(template_path, output_path, df_perusahaan, df_piutang):
    """
    Membuat laporan Identifier vs UMKM dengan sheet Total, NonPerform, dan NPF %.
    
    Args:
        template_path (str): Path template Excel
        output_path (str): Path output Excel
        df_perusahaan_path (str): Path Data Perusahaan.xlsx
        df_piutang_path (str): Path Data Piutang dan CKPN.xlsx
    """
    # Konfigurasi
    START_ROW, START_COL, END_COL = 3, 4, 9
    NUM_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    PERCENT_FORMAT = '0.00%'
    
    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL = PatternFill("solid", fgColor="FF99FF")
    BORDER_THIN = Border(left=Side(style="thin"), right=Side(style="thin"), 
                        top=Side(style="thin"), bottom=Side(style="thin"))
    
    UMKM_MAP = {"EN:e16": 4, "EN:e17": 5, "EN:e18": 6, "EN:e19": 7, "EN:e20": 8}
    
    def apply_border(ws, last_row):
        for r in range(2, last_row + 1):
            for c in range(START_COL, END_COL + 1):
                ws.cell(r, c).border = BORDER_THIN
    
    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]
            
            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH
            
            row += 1
        return row - 1
    
    def apply_total(ws, last_row, number_format):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)
        
        for c in range(1, END_COL + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL
        
        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = number_format
        
        return total_row
    
    def fill_value_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))
        for _, r in df_src.iterrows():
            if r["KTGR_SH_DBTR"] in UMKM_MAP:
                rekap[r["IDENTIFIER"]][r["KTGR_SH_DBTR"]] += r["PIUTANG"]
        
        last_row = fill_base(ws)
        
        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value
            for k, c in UMKM_MAP.items():
                cell = ws.cell(r, c)
                cell.value = rekap.get(ident, {}).get(k, 0)
                cell.number_format = NUM_FORMAT
            
            ws[f"I{r}"].value = f"=SUM(D{r}:H{r})"
            ws[f"I{r}"].number_format = NUM_FORMAT
        
        total_row = apply_total(ws, last_row, NUM_FORMAT)
        apply_border(ws, last_row)
        return last_row, total_row
    
    # Load template
    wb = load_workbook(template_path)
    
    # Total sheet
    fill_value_sheet(wb["Total"], df_piutang)
    
    # NonPerform sheet
    df_np = df_piutang[df_piutang["KLTS"].isin(["SF:e9", "SF:e3", "SF:e4"])]
    fill_value_sheet(wb["NonPerform"], df_np)
    
    # NPF sheet
    ws_npf = wb["NPF"]
    last_row = fill_base(ws_npf)
    
    for r in range(START_ROW, last_row + 1):
        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            ws_npf.cell(r, c).value = f"=IFERROR(NonPerform!{col}{r}/Total!{col}{r},0)"
            ws_npf.cell(r, c).number_format = PERCENT_FORMAT
    
    total_row = last_row + 2
    ws_npf[f"C{total_row}"] = "Total"
    ws_npf[f"C{total_row}"].font = Font(bold=True)
    
    for c in range(1, END_COL + 1):
        ws_npf.cell(total_row, c).fill = FILL_TOTAL
    
    for c in range(START_COL, END_COL + 1):
        col = get_column_letter(c)
        cell = ws_npf.cell(total_row, c)
        cell.value = f"=IFERROR(NonPerform!{col}{total_row}/Total!{col}{total_row},0)"
        cell.font = Font(bold=True)
        cell.number_format = PERCENT_FORMAT
    
    apply_border(ws_npf, last_row)
    
    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 5. UMKM vs. Lokasi
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook

def laporan_umkm_vs_lokasi(template_path, output_path, df_piutang):
    """
    Membuat laporan UMKM vs Lokasi berdasarkan FORM (PS, OK, SUS, PUP).
    
    Args:
        template_path (str): Path template Excel
        output_path (str): Path output Excel  
        df_piutang (str): Data Piutang dan CKPN.xlsx
    """
    # Konfigurasi
    START_ROW = 3
    
    UMKM_COL_MAP = {
        "EN:e16": "C",
        "EN:e17": "D", 
        "EN:e18": "E",
        "EN:e19": "F",
        "EN:e20": "G"
    }
    
    FORM_SHEET_MAP = {
        2110: "PS",
        2120: "OK", 
        2130: "SUS",
        2140: "PUP"
    }
    
    
    # Rekap data (FORM × LOKASI × UMKM)
    rekap = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    
    for _, r in df_piutang.iterrows():
        form = r.get("FORM")
        lokasi = r.get("LKS_DT__PRYK")
        umkm = r.get("KTGR_SH_DBTR")
        nilai = r.get("PIUTANG", 0)
        
        if (
            form in FORM_SHEET_MAP
            and pd.notna(lokasi)
            and umkm in UMKM_COL_MAP
        ):
            rekap[form][str(lokasi).strip()][umkm] += nilai
    
    # Load template
    wb = load_workbook(template_path)
    
    # Isi setiap sheet berdasarkan FORM
    for form, sheet_name in FORM_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        data_form = rekap.get(form, {})
        
        for row in range(START_ROW, ws.max_row + 1):
            kode_lokasi = ws[f"B{row}"].value
            if kode_lokasi is None:
                continue
            
            kode_lokasi = str(kode_lokasi).strip()
            data_lokasi = data_form.get(kode_lokasi, {})
            
            for umkm, col in UMKM_COL_MAP.items():
                cell = ws[f"{col}{row}"]
                
                # 🚫 JANGAN TIMPA FORMULA
                if cell.data_type == "f":
                    continue
                
                # ✅ Isi nilai / 0
                cell.value = data_lokasi.get(umkm, 0)
    
    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
### 6. UMKM vs. Lokasi (Kontrak)
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook

def laporan_umkm_vs_lokasi_kontrak(template_path, output_path, df_piutang):
    """
    Membuat laporan UMKM vs Lokasi berdasarkan FORM (PS, OK, SUS, PUP).
    
    Args:
        template_path (str): Path template Excel
        output_path (str): Path output Excel  
        df_piutang (str): Data Piutang dan CKPN.xlsx
    """
    # Konfigurasi
    START_ROW = 3
    
    UMKM_COL_MAP = {
        "EN:e16": "C",
        "EN:e17": "D", 
        "EN:e18": "E",
        "EN:e19": "F",
        "EN:e20": "G"
    }
    
    FORM_SHEET_MAP = {
        2110: "PS",
        2120: "OK", 
        2130: "SUS",
        2140: "PUP"
    }
    
    
    # Rekap data (FORM × LOKASI × UMKM)
    rekap = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    
    for _, r in df_piutang.iterrows():
        form = r.get("FORM")
        lokasi = r.get("LKS_DT__PRYK")
        umkm = r.get("KTGR_SH_DBTR")
        nilai = r.get("KONTRAK", 0)
        
        if (
            form in FORM_SHEET_MAP
            and pd.notna(lokasi)
            and umkm in UMKM_COL_MAP
        ):
            rekap[form][str(lokasi).strip()][umkm] += nilai
    
    # Load template
    wb = load_workbook(template_path)
    
    # Isi setiap sheet berdasarkan FORM
    for form, sheet_name in FORM_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        data_form = rekap.get(form, {})
        
        for row in range(START_ROW, ws.max_row + 1):
            kode_lokasi = ws[f"B{row}"].value
            if kode_lokasi is None:
                continue
            
            kode_lokasi = str(kode_lokasi).strip()
            data_lokasi = data_form.get(kode_lokasi, {})
            
            for umkm, col in UMKM_COL_MAP.items():
                cell = ws[f"{col}{row}"]
                
                # 🚫 JANGAN TIMPA FORMULA
                if cell.data_type == "f":
                    continue
                
                # ✅ Isi nilai / 0
                cell.value = data_lokasi.get(umkm, 0)
    
    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
### 7. UMKM vs. Sektor Ekonomi
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook

def laporan_umkm_vs_sektor_ekonomi(template_path, output_path, df_piutang):
    """
    Membuat laporan UMKM vs Sektor Ekonomi berdasarkan FORM (PS, OK, SUS, PUP).
    
    Args:
        template_path (str): Path template Excel
        output_path (str): Path output Excel
        df_piutang_path (str): Path Data Piutang dan CKPN.xlsx
    """
    # Konfigurasi
    START_ROW = 3
    
    UMKM_COL_MAP = {
        "EN:e16": "C",
        "EN:e17": "D",
        "EN:e18": "E", 
        "EN:e19": "F",
        "EN:e20": "G"
    }
    
    FORM_SHEET_MAP = {
        2110: "PS",
        2120: "OK",
        2130: "SUS",
        2140: "PUP"
    }

    
    # Rekap data (FORM × SEKTOR × UMKM)
    rekap = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    
    for _, r in df_piutang.iterrows():
        form = r.get("FORM")
        sektor = r.get("SKTR_KNM_LPNGN_SH")
        umkm = r.get("KTGR_SH_DBTR")
        nilai = r.get("PIUTANG", 0)
        
        if (
            form in FORM_SHEET_MAP
            and pd.notna(sektor)
            and umkm in UMKM_COL_MAP
        ):
            rekap[form][str(sektor).strip()][umkm] += nilai
    
    # Load template
    wb = load_workbook(template_path)
    
    # Isi setiap sheet berdasarkan FORM
    for form, sheet_name in FORM_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        data_form = rekap.get(form, {})
        
        for row in range(START_ROW, ws.max_row + 1):
            kode_sektor = ws[f"B{row}"].value
            if kode_sektor is None:
                continue
            
            kode_sektor = str(kode_sektor).strip()
            data_sektor = data_form.get(kode_sektor, {})
            
            for umkm, col in UMKM_COL_MAP.items():
                cell = ws[f"{col}{row}"]
                
                # 🚫 JANGAN TIMPA FORMULA
                if cell.data_type == "f":
                    continue
                
                # ✅ Isi nilai / 0
                cell.value = data_sektor.get(umkm, 0)
    
    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

## Jenis Pembiayaan
def laporan_jenis_pembiayaan_vs_npf(template_path: str, output_path: str, df_piutang: pd.DataFrame) -> None:
    """
    Mengisi data Jenis Pembiayaan vs NPF ke template Excel.
    
    Args:
        template_path (str): Path file template Excel
        output_path (str): Path file output Excel
        df_piutang (pd.DataFrame): DataFrame berisi data piutang dan CKPN
    """
    
    # =====================================================
    # 1️⃣ KONFIGURASI
    # =====================================================
    START_ROW = 3
    PIUTANG_COL = {
        "SF:e2": 3,
        "SF:e12": 4,
        "SF:e9": 5,
        "SF:e3": 6,
        "SF:e4": 7
    }
    CKPN_COL = {
        "SF:e2": 10,
        "SF:e12": 11,
        "SF:e9": 12,
        "SF:e3": 13,
        "SF:e4": 14
    }
    
    # =====================================================
    # 2️⃣ REKAP DATA
    # =====================================================
    rekap_piu = defaultdict(lambda: defaultdict(float))
    rekap_ckpn = defaultdict(lambda: defaultdict(float))
    
    for _, r in df_piutang.iterrows():
        sandi = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
        klts = r.get("KLTS")
        
        if pd.notna(sandi) and klts in PIUTANG_COL:
            rekap_piu[sandi][klts] += r.get("PIUTANG", 0)
            rekap_ckpn[sandi][klts] += r.get("CKPN", 0)
    
    # =====================================================
    # 3️⃣ LOAD & ISI TEMPLATE
    # =====================================================
    wb = load_workbook(template_path)
    ws = wb.active
    
    for row in range(START_ROW, ws.max_row + 1):
        sandi = ws.cell(row, 2).value  # kolom B
        if not sandi:
            continue
        
        sandi = str(sandi).strip()
        
        # ---- PIUTANG ----
        for klts, col in PIUTANG_COL.items():
            cell = ws.cell(row, col)
            if cell.data_type == "f":
                continue
            cell.value = rekap_piu.get(sandi, {}).get(klts, 0)
        
        # ---- CKPN ----
        for klts, col in CKPN_COL.items():
            cell = ws.cell(row, col)
            if cell.data_type == "f":
                continue
            cell.value = rekap_ckpn.get(sandi, {}).get(klts, 0)
    
    # =====================================================
    # 4️⃣ SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
## Piutang Valas
def laporan_piutang_valas(template_path, output_path, df_perusahaan, df_piutang):

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW, START_COL, END_COL = 2, 4, 8
    NUM_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL   = PatternFill("solid", fgColor="FF99FF")
    BORDER_THIN  = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    MU_MAP = {
        "MU:IDR": 4,
        "MU:SGD": 5,
        "MU:USD": 6,
        "MU:MYR": 7
    }

    FORM_SHEET_MAP = {
        2110: "PS",
        2120: "OK",
        2130: "SUS",
        2140: "PUP"
    }

    # =====================================================
    # HELPER
    # =====================================================
    def apply_border(ws, last_row):
        for r in range(2, last_row + 1):
            for c in range(START_COL, END_COL + 1):
                ws.cell(r, c).border = BORDER_THIN

    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]

            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH
            row += 1
        return row - 1

    def apply_total(ws, last_row):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)

        for c in range(1, END_COL + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL

        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = NUM_FORMAT

        return total_row

    def fill_value_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))

        for _, r in df_src.iterrows():
            if r["JNS_MT_NG"] in MU_MAP:
                rekap[r["IDENTIFIER"]][r["JNS_MT_NG"]] += r["PIUTANG"]

        last_row = fill_base(ws)

        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value

            for k, c in MU_MAP.items():
                cell = ws.cell(r, c)
                cell.value = rekap.get(ident, {}).get(k, 0)
                cell.number_format = NUM_FORMAT

            ws[f"H{r}"].value = f"=SUM(D{r}:G{r})"
            ws[f"H{r}"].number_format = NUM_FORMAT

        apply_total(ws, last_row)
        apply_border(ws, last_row)

    # =====================================================
    # LOAD TEMPLATE
    # =====================================================
    wb = load_workbook(template_path)

    # =====================================================
    # LOOP FORM → SHEET
    # =====================================================
    for form, sheet_name in FORM_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        df_form = df_piutang[df_piutang["FORM"] == form]

        if df_form.empty:
            continue

        fill_value_sheet(ws, df_form)
    
    # ===============================
    # PIUTANG VALAS (GABUNGAN)
    # ===============================
    ws_valas = wb["Piutang Valas"]
    last_row = fill_base(ws_valas)

    for r in range(START_ROW, last_row + 1):
        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            ws_valas[f"{col}{r}"].value = (
                f"=PS!{col}{r}+OK!{col}{r}+SUS!{col}{r}+PUP!{col}{r}"
            )
            ws_valas[f"{col}{r}"].number_format = NUM_FORMAT

        ws_valas[f"H{r}"].value = f"=SUM(D{r}:G{r})"
        ws_valas[f"H{r}"].number_format = NUM_FORMAT
    
    apply_total(ws_valas, last_row)
    apply_border(ws_valas, last_row)

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
## Pinjaman Valas
def laporan_pinjaman_valas(template_path, output_path, df_perusahaan, df_pinjaman_valas):

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 2
    START_COL = 4      # D
    DATA_END_COL = 6   # F
    TOTAL_COL = 7     # G

    NUM_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL   = PatternFill("solid", fgColor="FF99FF")
    BORDER_THIN  = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    MU_MAP = {
        "MU:EUR": 4,   # D
        "MU:IDR": 5,   # E
        "MU:USD": 6    # F
    }

    # =====================================================
    # HELPER
    # =====================================================
    def apply_border(ws, last_row):
        for r in range(2, last_row + 1):
            for c in range(START_COL, TOTAL_COL + 1):
                ws.cell(r, c).border = BORDER_THIN

    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]

            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH

            row += 1
        return row - 1

    def apply_total(ws, last_row):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)

        for c in range(1, TOTAL_COL + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL

        for c in range(START_COL, DATA_END_COL + 1):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = NUM_FORMAT

        ws.cell(total_row, TOTAL_COL).value = f"=SUM(D{total_row}:F{total_row})"
        ws.cell(total_row, TOTAL_COL).number_format = NUM_FORMAT

    def fill_value_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))

        for _, r in df_src.iterrows():
            mu = r.get("JNS_MT_NG")
            if mu in MU_MAP:
                rekap[r["IDENTIFIER"]][mu] += r.get("SLD_PNJMN", 0)

        last_row = fill_base(ws)

        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value
            if not ident:
                continue

            for mu, c in MU_MAP.items():
                cell = ws.cell(r, c)
                if cell.data_type == "f":
                    continue
                cell.value = rekap.get(ident, {}).get(mu, 0)
                cell.number_format = NUM_FORMAT

            ws.cell(r, TOTAL_COL).value = f"=SUM(D{r}:F{r})"
            ws.cell(r, TOTAL_COL).number_format = NUM_FORMAT

        apply_total(ws, last_row)
        apply_border(ws, last_row)

    # =====================================================
    # LOAD & EXECUTE
    # =====================================================
    wb = load_workbook(template_path)
    ws = wb.active   # hanya 1 sheet

    fill_value_sheet(ws, df_pinjaman_valas)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
## Jumlah Kontrak
### 1. Identifier vs. Jenis Usaha
def laporan_identifier_vs_jenis_usaha_kontrak(template_path, output_path, df_perusahaan, df_piutang):

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 3
    START_COL = 4      # D
    END_COL   = 14     # N

    NUM_FORMAT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'

    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL   = PatternFill("solid", fgColor="FF99FF")
    BORDER_THIN  = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    USAHA_COL = {
        "AK:e20": 5,   # E
        "AK:e21": 6,   # F
        "AK:e22": 7,   # G
        "AK:e23": 8,   # H
        "AK:e24": 10,  # J
        "AK:e25": 11,  # K
        "AK:e26": 12,  # L
        "AK:e27": 13   # M
    }

    # =====================================================
    # HELPER
    # =====================================================
    def apply_border(ws, last_row):
        for r in range(2, last_row + 1):
            for c in range(START_COL, END_COL + 1):
                ws.cell(r, c).border = BORDER_THIN

    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]

            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH
            row += 1
        return row - 1

    def apply_total(ws, last_row):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)

        for c in range(1, END_COL + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL

        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = NUM_FORMAT

    def fill_value_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))

        for _, r in df_src.iterrows():
            usaha = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
            if usaha in USAHA_COL:
                rekap[r["IDENTIFIER"]][usaha] += r.get("KONTRAK", 0)

        last_row = fill_base(ws)

        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value
            if not ident:
                continue

            # --- ISI NILAI USAHA ---
            for usaha, c in USAHA_COL.items():
                cell = ws.cell(r, c)
                if cell.data_type == "f":
                    continue
                cell.value = rekap.get(ident, {}).get(usaha, 0)
                cell.number_format = NUM_FORMAT

            # --- RUMUS TAMBAHAN ---
            ws[f"D{r}"].value = f"=SUM(E{r}:H{r})"
            ws[f"I{r}"].value = f"=SUM(J{r}:M{r})"
            ws[f"N{r}"].value = f"=D{r}+I{r}"

            for col in ["D", "I", "N"]:
                ws[f"{col}{r}"].number_format = NUM_FORMAT

        apply_total(ws, last_row)
        apply_border(ws, last_row)

    # =====================================================
    # LOAD & EXECUTE
    # =====================================================
    wb = load_workbook(template_path)
    ws = wb.active

    fill_value_sheet(ws, df_piutang)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
### 2. Identifier vs. Kolektibilitas
def laporan_identifier_vs_kolektibilitas(
    template_path,
    output_path,
    df_perusahaan,
    df_piutang
):

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW, START_COL, END_COL = 3, 4, 10   # D–H
    NUM_FORMAT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'

    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL   = PatternFill("solid", fgColor="FF99FF")
    BORDER_THIN  = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # KOLEKTABILITAS → KOLOM
    KOLEK_COL = {
        "SF:e2": 4,   # D
        "SF:e12": 5,  # E
        "SF:e9": 6,   # F
        "SF:e3": 7,   # G
        "SF:e4": 8    # H
    }

    FORM_SHEET_MAP = {
        2110: "PS",
        2120: "OK",
        2130: "SUS",
        2140: "PUP"
    }

    # =====================================================
    # HELPER
    # =====================================================
    def apply_border(ws, last_row):
        for r in range(2, last_row + 1):
            for c in range(START_COL, 11):  # sampai J
                ws.cell(r, c).border = BORDER_THIN

    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]

            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH
            row += 1
        return row - 1

    def apply_total(ws, last_row):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)

        for c in range(1, 11):
            ws.cell(total_row, c).fill = FILL_TOTAL

        for c in range(START_COL, 11):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = NUM_FORMAT

    def fill_value_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))

        for _, r in df_src.iterrows():
            klts = r.get("KLTS")
            if klts in KOLEK_COL:
                rekap[r["IDENTIFIER"]][klts] += r.get("KONTRAK", 0)

        last_row = fill_base(ws)

        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value
            if not ident:
                continue

            # D–H
            for klts, c in KOLEK_COL.items():
                cell = ws.cell(r, c)
                if cell.data_type == "f":
                    continue
                cell.value = rekap.get(ident, {}).get(klts, 0)
                cell.number_format = NUM_FORMAT

            # I = SUM(F:H)
            ws[f"I{r}"].value = f"=SUM(F{r}:H{r})"
            ws[f"I{r}"].number_format = NUM_FORMAT

            # J = SUM(D:H)
            ws[f"J{r}"].value = f"=SUM(D{r}:H{r})"
            ws[f"J{r}"].number_format = NUM_FORMAT

        apply_total(ws, last_row)
        apply_border(ws, last_row)

    # =====================================================
    # LOAD TEMPLATE
    # =====================================================
    wb = load_workbook(template_path, keep_links=False)

    # =====================================================
    # ISI SHEET PS / OK / SUS / PUP
    # =====================================================
    for form, sheet_name in FORM_SHEET_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        df_form = df_piutang[df_piutang["FORM"] == form]

        if not df_form.empty:
            fill_value_sheet(ws, df_form)

    # =====================================================
    # SHEET FIX (GABUNGAN)
    # =====================================================
    ws_fix = wb["FIX"]
    last_row = fill_base(ws_fix)

    for r in range(START_ROW, last_row + 1):

        # D–H = PS + OK + SUS + PUP
        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            ws_fix[f"{col}{r}"].value = (
                f"=PS!{col}{r}+OK!{col}{r}+SUS!{col}{r}+PUP!{col}{r}"
            )
            ws_fix[f"{col}{r}"].number_format = NUM_FORMAT

        ws_fix[f"I{r}"].value = f"=SUM(F{r}:H{r})"
        ws_fix[f"I{r}"].number_format = NUM_FORMAT

        ws_fix[f"J{r}"].value = f"=SUM(D{r}:H{r})"
        ws_fix[f"J{r}"].number_format = NUM_FORMAT

    apply_total(ws_fix, last_row)
    apply_border(ws_fix, last_row)

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
### 3. Jenis Usaha vs. Kolektibilitas
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook


def laporan_jenis_pembiayaan_vs_kolektibilitas(template_path: str, output_path: str, df_piutang: pd.DataFrame) -> None:
    """
    Mengisi data Jenis Pembiayaan vs NPF ke template Excel.
    
    Args:
        template_path (str): Path file template Excel
        output_path (str): Path file output Excel
        df_piutang (pd.DataFrame): DataFrame berisi data piutang dan CKPN
    """
    
    # =====================================================
    # 1️⃣ KONFIGURASI
    # =====================================================
    START_ROW = 3
    KLTS_COL = {
        "SF:e2": 3,
        "SF:e12": 4,
        "SF:e9": 5,
        "SF:e3": 6,
        "SF:e4": 7

    }
    
    # =====================================================
    # 2️⃣ REKAP DATA
    # =====================================================
    rekap_kol = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        sandi = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
        klts = r.get("KLTS")
        
        if pd.notna(sandi) and klts in KLTS_COL:
            rekap_kol[sandi][klts] += r.get("KONTRAK", 0)

    
    # =====================================================
    # 3️⃣ LOAD & ISI TEMPLATE
    # =====================================================
    wb = load_workbook(template_path)
    ws = wb.active
    
    for row in range(START_ROW, ws.max_row + 1):
        sandi = ws.cell(row, 2).value  # kolom B
        if not sandi:
            continue
        
        sandi = str(sandi).strip()
        
        # ---- PIUTANG ----
        for klts, col in KLTS_COL.items():
            cell = ws.cell(row, col)
            if cell.data_type == "f":
                continue
            cell.value = rekap_kol.get(sandi, {}).get(klts, 0)
        
    
    # =====================================================
    # 4️⃣ SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
### 4. Lokasi vs. Jenis Usaha
def laporan_lokasi_vs_jenis_usaha_kontrak(template_path, output_path, df_piutang):
    """
    Mengisi laporan Jenis Usaha vs Lokasi
    - Sandi lokasi: kolom B
    - Jenis usaha → kolom sesuai mapping
    - Tidak menimpa formula
    """

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 3

    USAHA_COL = {
        "AK:e20": 4,   # D
        "AK:e21": 5,   # E
        "AK:e22": 6,   # F
        "AK:e23": 7,   # G
        "AK:e24": 9,   # I
        "AK:e25": 10,  # J
        "AK:e26": 11,  # K
        "AK:e27": 12   # L
    }

    # =====================================================
    # REKAP DATA (LOKASI × JENIS USAHA)
    # =====================================================
    rekap = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        lokasi = r.get("LKS_DT__PRYK")
        usaha  = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
        nilai  = r.get("KONTRAK", 0)

        if pd.notna(lokasi) and usaha in USAHA_COL:
            rekap[str(lokasi).strip()][usaha] += nilai

    # =====================================================
    # LOAD TEMPLATE
    # =====================================================
    wb = load_workbook(template_path)
    ws = wb.active   # hanya 1 sheet

    # =====================================================
    # ISI KE TEMPLATE
    # =====================================================
    for row in range(START_ROW, ws.max_row + 1):

        lokasi = ws.cell(row, 2).value   # kolom B
        if not lokasi:
            continue

        lokasi = str(lokasi).strip()
        data_lokasi = rekap.get(lokasi, {})

        for usaha, col in USAHA_COL.items():
            cell = ws.cell(row, col)

            # 🚫 Jangan timpa formula
            if cell.data_type == "f":
                continue

            cell.value = data_lokasi.get(usaha, 0)

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 5. Sektor Ekonomi vs. Jenis Usaha
def laporan_sektor_vs_jenis_usaha_kontrak(template_path, output_path, df_piutang):

    wb = load_workbook(template_path)
    ws = wb["FIX"]

    MAP_AKUN_COL = {
        "AK:e20": "D",
        "AK:e21": "E",
        "AK:e22": "F",
        "AK:e23": "G",
        "AK:e24": "I",
        "AK:e25": "J",
        "AK:e26": "K",
        "AK:e27": "L"
    }

    rekap = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        sektor = r["SKTR_KNM_LPNGN_SH"]
        akun   = r["JNS_KGTN_SH_PRSHN_MDL_VNTR"]

        if akun in MAP_AKUN_COL:
            rekap[sektor][akun] += r["KONTRAK"]

    for row in range(1, ws.max_row + 1):
        sektor = ws[f"B{row}"].value
        if sektor not in rekap:
            continue

        for akun, col in MAP_AKUN_COL.items():
            if akun in rekap[sektor]:
                ws[f"{col}{row}"] = rekap[sektor][akun]

    wb.save(output_path)

## Lokasi
### 1. Lokasi vs. NPF
def laporan_lokasi_vs_npf(template_path, output_path, df_piutang):

    wb = load_workbook(template_path)

    MAP_KLTS_PIUTANG = {
        "SF:e2":  "C",
        "SF:e12": "D",
        "SF:e9":  "E",
        "SF:e3":  "F",
        "SF:e4":  "G"
    }

    MAP_KLTS_CKPN = {
        "SF:e2":  "J",
        "SF:e12": "K",
        "SF:e9":  "L",
        "SF:e3":  "M",
        "SF:e4":  "N"
    }

    FORM_TO_SHEET = {
    "2110": "PS",
    "2120": "OK",
    "2130": "SUS",
    "2140": "PUP"
    }   

    rekap_piutang = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    rekap_ckpn    = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for _, r in df_piutang.iterrows():
        form   = str(r["FORM"])
        sektor = r["LKS_DT__PRYK"]
        klts   = r["KLTS"]

        if pd.notna(r.get("PIUTANG")):
            rekap_piutang[form][sektor][klts] += r["PIUTANG"]

        if pd.notna(r.get("CKPN")):
            rekap_ckpn[form][sektor][klts] += r["CKPN"]

    for form, sheet in FORM_TO_SHEET.items():
        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]

        for row in range(3, ws.max_row + 1):
            sektor = ws[f"B{row}"].value
            if not sektor:
                continue

            data_piu = rekap_piutang.get(form, {}).get(sektor,0)
            data_ckp = rekap_ckpn.get(form, {}).get(sektor,0)

            if not data_piu:
                continue

            for klts, col in MAP_KLTS_PIUTANG.items():
                if klts in data_piu:
                    ws[f"{col}{row}"] = data_piu[klts]

            if sheet != "PS" and data_ckp:
                for klts, col in MAP_KLTS_CKPN.items():
                    if klts in data_ckp:
                        ws[f"{col}{row}"] = data_ckp[klts]

    wb.save(output_path)

### 2. Lokasi vs. Jenis Usaha (NPF)
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook

def laporan_lokasi_vs_jenis_usaha_npf(template_path, output_path, df_piutang):
    """
    Mengisi laporan Lokasi vs Jenis Usaha
    Sheet:
    - Total 34 Prov
    - NonPerform 34 Prov
    - Total 38 Prov
    - NonPerform 38 Prov
    """

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 3

    USAHA_COL = {
        "AK:e20": 4,   # D
        "AK:e21": 5,   # E
        "AK:e22": 6,   # F
        "AK:e23": 7,   # G
        "AK:e24": 9,   # I
        "AK:e25": 10,  # J
        "AK:e26": 11,  # K
        "AK:e27": 12   # L
    }

    SHEET_CONFIG = {
        "Total 34 Prov": df_piutang,
        "Total 38 Prov": df_piutang,
        "NonPerform 34 Prov": df_piutang[
            df_piutang["KLTS"].isin(["SF:e9", "SF:e3", "SF:e4"])
        ],
        "NonPerform 38 Prov": df_piutang[
            df_piutang["KLTS"].isin(["SF:e9", "SF:e3", "SF:e4"])
        ],
    }

    # =====================================================
    # HELPER: ISI 1 SHEET
    # =====================================================
    def fill_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))

        for _, r in df_src.iterrows():
            lokasi = r.get("LKS_DT__PRYK")
            usaha  = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
            nilai  = r.get("PIUTANG", 0)

            if pd.notna(lokasi) and usaha in USAHA_COL:
                rekap[str(lokasi).strip()][usaha] += nilai

        for row in range(START_ROW, ws.max_row + 1):
            lokasi = ws.cell(row, 2).value  # kolom B
            if not lokasi:
                continue

            lokasi = str(lokasi).strip()
            data_lokasi = rekap.get(lokasi, {})

            for usaha, col in USAHA_COL.items():
                cell = ws.cell(row, col)

                # 🚫 Jangan timpa formula
                if cell.data_type == "f":
                    continue

                cell.value = data_lokasi.get(usaha, 0)

    # =====================================================
    # LOAD TEMPLATE & PROSES SEMUA SHEET
    # =====================================================
    wb = load_workbook(template_path)

    for sheet_name, df_src in SHEET_CONFIG.items():
        if sheet_name not in wb.sheetnames:
            continue

        fill_sheet(wb[sheet_name], df_src)

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 3. Lokasi vs. Identifier
def laporan_lokasi_vs_identifier(
    template_path,
    output_path,
    df_perusahaan,
    df_piutang
):
    """
    Mengisi laporan lokasi vs identifier untuk:
    - 34 Prov
    - 38 Prov
    (logika sama, jumlah perusahaan dinamis, TOTAL otomatis)
    Template mulai dari 1 kolom C, kolom baru otomatis copy formula+style
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment
    from copy import copy
    from collections import defaultdict
    from openpyxl.formula.translate import Translator
    import os

    # =====================================================
    # KONFIGURASI
    # =====================================================
    SHEET_NAMES = ["34 Prov", "38 Prov"]
    START_COL = 3        # Kolom C
    HEADER_ROW_1 = 1
    HEADER_ROW_2 = 2
    START_ROW_DATA = 3

    # Helper function untuk copy kolom lengkap (formula + style)
    def copy_col_with_style_and_formula(ws, src_col_idx, dst_col_idx, row_start=1, row_end=None):
        if row_end is None:
            row_end = ws.max_row
        
        for r in range(row_start, row_end + 1):
            src = ws.cell(row=r, column=src_col_idx)
            dst = ws.cell(row=r, column=dst_col_idx)

            # Copy style lengkap
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)

            # Copy value (kalau formula, translate referensi)
            v = src.value
            if isinstance(v, str) and v.startswith("="):
                # Formula otomatis adjust: C3=SUM(C4:C5) → D3=SUM(D4:D5)
                dst.value = Translator(v, origin=src.coordinate).translate_formula(dst.coordinate)
            else:
                dst.value = v

    # Load workbook
    wb = load_workbook(template_path)
    ident_list = df_perusahaan["IDENTIFIER"].tolist()

    # =====================================================
    # REKAP DATA (LOKASI × IDENTIFIER)
    # =====================================================
    rekap = defaultdict(lambda: defaultdict(float))
    for _, r in df_piutang.iterrows():
        rekap[r["LKS_DT__PRYK"]][r["IDENTIFIER"]] += r["PIUTANG"]

    # =====================================================
    # PROSES SETIAP SHEET
    # =====================================================
    for sheet_name in SHEET_NAMES:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        # -------------------------------------------------
        # 1️⃣ TAMBAH KOLOM PERUSAHAAN (tanpa hapus kolom)
        # -------------------------------------------------
        required_cols = len(ident_list)
        current_company_cols = ws.max_column - START_COL + 1  # Kolom perusahaan saat ini

        # Tambah kolom sampai cukup
        while current_company_cols < required_cols:
            insert_at = START_COL + current_company_cols  # Posisi kolom baru
            ws.insert_cols(insert_at)
            
            # Copy dari kolom sebelumnya (template kolom C akan propagate)
            src_col_idx = insert_at - 1
            dst_col_idx = insert_at
            copy_col_with_style_and_formula(ws, src_col_idx, dst_col_idx)
            
            current_company_cols += 1

        # -------------------------------------------------
        # 2️⃣ ISI HEADER IDENTIFIER & NAMA PERUSAHAAN
        # -------------------------------------------------
        for i, p in enumerate(df_perusahaan.itertuples(index=False)):
            col = get_column_letter(START_COL + i)
            
            # Header 1: IDENTIFIER → CENTER
            header1_cell = ws[f"{col}{HEADER_ROW_1}"]
            header1_cell.value = p.IDENTIFIER
            header1_cell.alignment = Alignment(horizontal='center')
            
            # Header 2: NAMA PERUSAHAAN → tetap original
            ws[f"{col}{HEADER_ROW_2}"] = p.NM_SBTN_PRSHN

        # -------------------------------------------------
        # 3️⃣ ISI DATA (hanya override VALUE data cells, formula aman)
        # -------------------------------------------------
        for row in range(START_ROW_DATA, ws.max_row + 1):
            lokasi = ws[f"B{row}"].value
            if not lokasi:
                continue

            for i, ident in enumerate(ident_list):
                col = get_column_letter(START_COL + i)
                cell = ws[f"{col}{row}"]
                
                # **CEK FORMULA DULU**
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue  # Formula → Skip
                
                # Data → Isi (0 jadi kosong/None)
                data_value = rekap[lokasi].get(ident, 0) if lokasi in rekap else 0
                cell.value = data_value

        # -------------------------------------------------
        # 4️⃣ BUAT KOLOM TOTAL
        # -------------------------------------------------
        total_col_idx = START_COL + len(ident_list)
        total_col = get_column_letter(total_col_idx)
        ws.insert_cols(total_col_idx)

        # Copy style dari kolom kiri terakhir
        last_data_col_idx = total_col_idx - 1
        for r in range(1, ws.max_row + 1):
            src = ws.cell(r, last_data_col_idx)
            dst = ws.cell(r, total_col_idx)
            
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format

        ws[f"{total_col}{HEADER_ROW_2}"] = "TOTAL"
        
        # Formula TOTAL per baris
        last_data_col_letter = get_column_letter(total_col_idx - 1)
        for r in range(START_ROW_DATA, ws.max_row + 1):
            ws[f"{total_col}{r}"] = f"=SUM(C{r}:{last_data_col_letter}{r})"

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()

### 4. Lokasi vs. Identifier (Kontrak)
def laporan_lokasi_vs_identifier_kontrak(
    template_path,
    output_path,
    df_perusahaan,
    df_piutang
):
    """
    Mengisi laporan lokasi vs identifier untuk:
    - 34 Prov
    - 38 Prov
    (logika sama, jumlah perusahaan dinamis, TOTAL otomatis)
    Template mulai dari 1 kolom C, kolom baru otomatis copy formula+style
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment
    from copy import copy
    from collections import defaultdict
    from openpyxl.formula.translate import Translator
    import os

    # =====================================================
    # KONFIGURASI
    # =====================================================
    SHEET_NAMES = ["34 Prov", "38 Prov"]
    START_COL = 3        # Kolom C
    HEADER_ROW_1 = 1
    HEADER_ROW_2 = 2
    START_ROW_DATA = 3

    # Helper function untuk copy kolom lengkap (formula + style)
    def copy_col_with_style_and_formula(ws, src_col_idx, dst_col_idx, row_start=1, row_end=None):
        if row_end is None:
            row_end = ws.max_row
        
        for r in range(row_start, row_end + 1):
            src = ws.cell(row=r, column=src_col_idx)
            dst = ws.cell(row=r, column=dst_col_idx)

            # Copy style lengkap
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)

            # Copy value (kalau formula, translate referensi)
            v = src.value
            if isinstance(v, str) and v.startswith("="):
                # Formula otomatis adjust: C3=SUM(C4:C5) → D3=SUM(D4:D5)
                dst.value = Translator(v, origin=src.coordinate).translate_formula(dst.coordinate)
            else:
                dst.value = v

    # Load workbook
    wb = load_workbook(template_path)
    ident_list = df_perusahaan["IDENTIFIER"].tolist()

    # =====================================================
    # REKAP DATA (LOKASI × IDENTIFIER)
    # =====================================================
    rekap = defaultdict(lambda: defaultdict(float))
    for _, r in df_piutang.iterrows():
        rekap[r["LKS_DT__PRYK"]][r["IDENTIFIER"]] += r["KONTRAK"]

    # =====================================================
    # PROSES SETIAP SHEET
    # =====================================================
    for sheet_name in SHEET_NAMES:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        # -------------------------------------------------
        # 1️⃣ TAMBAH KOLOM PERUSAHAAN (tanpa hapus kolom)
        # -------------------------------------------------
        required_cols = len(ident_list)
        current_company_cols = ws.max_column - START_COL + 1  # Kolom perusahaan saat ini

        # Tambah kolom sampai cukup
        while current_company_cols < required_cols:
            insert_at = START_COL + current_company_cols  # Posisi kolom baru
            ws.insert_cols(insert_at)
            
            # Copy dari kolom sebelumnya (template kolom C akan propagate)
            src_col_idx = insert_at - 1
            dst_col_idx = insert_at
            copy_col_with_style_and_formula(ws, src_col_idx, dst_col_idx)
            
            current_company_cols += 1

        # -------------------------------------------------
        # 2️⃣ ISI HEADER IDENTIFIER & NAMA PERUSAHAAN
        # -------------------------------------------------
        for i, p in enumerate(df_perusahaan.itertuples(index=False)):
            col = get_column_letter(START_COL + i)
            
            # Header 1: IDENTIFIER → CENTER
            header1_cell = ws[f"{col}{HEADER_ROW_1}"]
            header1_cell.value = p.IDENTIFIER
            header1_cell.alignment = Alignment(horizontal='center')
            
            # Header 2: NAMA PERUSAHAAN → tetap original
            ws[f"{col}{HEADER_ROW_2}"] = p.NM_SBTN_PRSHN

        # -------------------------------------------------
        # 3️⃣ ISI DATA (hanya override VALUE data cells, formula aman)
        # -------------------------------------------------
        for row in range(START_ROW_DATA, ws.max_row + 1):
            lokasi = ws[f"B{row}"].value
            if not lokasi:
                continue

            for i, ident in enumerate(ident_list):
                col = get_column_letter(START_COL + i)
                cell = ws[f"{col}{row}"]
                
                # **CEK FORMULA DULU**
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue  # Formula → Skip
                
                # Data → Isi (0 jadi kosong/None)
                data_value = rekap[lokasi].get(ident, 0) if lokasi in rekap else 0
                cell.value = data_value

        # -------------------------------------------------
        # 4️⃣ BUAT KOLOM TOTAL
        # -------------------------------------------------
        total_col_idx = START_COL + len(ident_list)
        total_col = get_column_letter(total_col_idx)
        ws.insert_cols(total_col_idx)

        # Copy style dari kolom kiri terakhir
        last_data_col_idx = total_col_idx - 1
        for r in range(1, ws.max_row + 1):
            src = ws.cell(r, last_data_col_idx)
            dst = ws.cell(r, total_col_idx)
            
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format

        ws[f"{total_col}{HEADER_ROW_2}"] = "TOTAL"
        
        # Formula TOTAL per baris
        last_data_col_letter = get_column_letter(total_col_idx - 1)
        for r in range(START_ROW_DATA, ws.max_row + 1):
            ws[f"{total_col}{r}"] = f"=SUM(C{r}:{last_data_col_letter}{r})"

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()

### 5. Lokasi vs. Sektor Ekonomi
def laporan_lokasi_vs_sektor(
    template_path,
    output_path,
    df_piutang
):
    from collections import defaultdict
    from openpyxl import load_workbook
    import pandas as pd
    import os

    # =====================================================
    # KONFIGURASI
    # =====================================================
    START_ROW = 3
    START_COL = 3      # C
    END_COL   = 26     # Z

    NP_KLTS = ["SF:e9", "SF:e3", "SF:e4"]

    FORM_MAP = {
        "PS": 2110,
        "OK": 2120,
        "SUS": 2130,
        "PUP": 2140
    }

    # =====================================================
    # LOAD TEMPLATE
    # =====================================================
    wb = load_workbook(template_path)

    # =====================================================
    # UTILITAS ISI DATA SAJA
    # =====================================================
    def fill_sheet(ws, df_src):
        """
        Mengisi nilai saja, tanpa:
        - ubah formula
        - ubah style
        """
        # Header sektor → kolom
        sektor_col = {
            ws.cell(1, c).value: c
            for c in range(START_COL, END_COL + 1)
            if ws.cell(1, c).value
        }

        # Rekap: lokasi × sektor
        rekap = defaultdict(lambda: defaultdict(float))
        for _, r in df_src.iterrows():
            lokasi = r["LKS_DT__PRYK"]
            sektor = r["SEKTOR_KAT"]
            nilai  = r["PIUTANG"]

            if pd.notna(lokasi) and pd.notna(sektor):
                rekap[str(lokasi).strip()][sektor] += nilai

        # Isi ke template
        for row in range(START_ROW, ws.max_row + 1):
            lokasi = ws.cell(row, 2).value  # kolom B
            if not lokasi:
                continue

            lokasi = str(lokasi).strip()
            data_lokasi = rekap.get(lokasi, {})

            for sektor, col in sektor_col.items():
                cell = ws.cell(row, col)

                # 🚫 jangan sentuh formula
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue

                cell.value = data_lokasi.get(sektor, 0)

    # =====================================================
    # PROSES PER FORM
    # =====================================================
    for label, form in FORM_MAP.items():

        # -------- TOTAL --------
        df_form = df_piutang[df_piutang["FORM"] == form]
        fill_sheet(wb[f"Total {label}"], df_form)

        # -------- NON PERFORM --------
        df_np = df_form[df_form["KLTS"].isin(NP_KLTS)]
        fill_sheet(wb[f"NP {label}"], df_np)

    # =====================================================
    # SAVE
    # =====================================================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

## KUKB
### 1. KUKB vs. NPF
def laporan_kukb_vs_npf(template_path, output_path, df_piutang):
    START_ROW = 3

    PIUTANG_COL = {
        "SF:e2": 2,
        "SF:e12": 3,
        "SF:e9": 4,
        "SF:e3": 5,
        "SF:e4": 6
    }

    CKPN_COL = {
        "SF:e2": 10,
        "SF:e12": 11,
        "SF:e9": 12,
        "SF:e3": 13,
        "SF:e4": 14
    }

    wb = load_workbook(template_path)
 
    ws = wb.active

    rekap_piu  = defaultdict(lambda: defaultdict(float))
    rekap_ckpn = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        ktgr = r.get("KTGR_SH_KNGN_BRKLNJTN")
        klts = r.get("KLTS")

        if pd.notna(ktgr) and klts in PIUTANG_COL:
            rekap_piu[ktgr][klts]  += r.get("PIUTANG", 0)
            rekap_ckpn[ktgr][klts] += r.get("CKPN", 0)

    row = START_ROW
    while ws.cell(row, 2).value is not None:

        sandi = ws.cell(row, 2).value
        ws.cell(row, 1).value = sandi
        ws.cell(row, 2).value = sandi

        for k, c in PIUTANG_COL.items():
            ws.cell(row, c).value = rekap_piu.get(sandi, {}).get(k, 0)

        for k, c in CKPN_COL.items():
            ws.cell(row, c).value = rekap_ckpn.get(sandi, {}).get(k, 0)

        row += 1

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
### 2. KUKB vs. Pariwisata
def laporan_kukb_vs_pariwisata(template_path, output_path, df_piutang):
    from collections import defaultdict
    from openpyxl import load_workbook
    import os

    wb = load_workbook(template_path)
    ws = wb.active

    START_ROW = 3  # Mulai isi data dari baris 3

    MAP_AKUN_COL = {
        "EN:e3819": "C", "EN:e3820": "D", "EN:e3821": "E", "EN:e3822": "F",
        "EN:e3823": "G", "EN:e3824": "H", "EN:e3825": "I", "EN:e3826": "J",
        "EN:e3827": "K", "EN:e3828": "L", "EN:e3829": "M", "EN:e3830": "N",
        "EN:e3831": "O"
    }

    rekap = defaultdict(lambda: defaultdict(float))

    # Rekap data
    for _, r in df_piutang.iterrows():
        sektor = r["SKTR_KNM_LPNGN_SH"]
        akun   = r["KTGR_SH_KNGN_BRKLNJTN"]
        if akun in MAP_AKUN_COL:
            rekap[sektor][akun] += r["PIUTANG"]

    # Proses baris 3+ - baris 1&2 aman, formula aman
    for row in range(START_ROW, ws.max_row + 1):
        sektor = ws[f"B{row}"].value
        
        for akun, col in MAP_AKUN_COL.items():
            cell = ws[f"{col}{row}"]
            
            # **CEK FORMULA - JANGAN SENTUH!**
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # Formula → Skip
            
            # Data → Isi 0/data
            value = rekap.get(sektor, {}).get(akun, 0)
            cell.value = value

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()
### 3. Ekonomi vs. Ekonomi Kreatif
def laporan_kukb_vs_ekraf(template_path, output_path, df_piutang):
    from collections import defaultdict
    from openpyxl import load_workbook
    import os

    wb = load_workbook(template_path)
    ws = wb.active

    START_ROW = 3  # Mulai isi data dari baris 3

    MAP_AKUN_COL = {
        "EN:e3819": "C", "EN:e3820": "D", "EN:e3821": "E", "EN:e3822": "F",
        "EN:e3823": "G", "EN:e3824": "H", "EN:e3825": "I", "EN:e3826": "J",
        "EN:e3827": "K", "EN:e3828": "L", "EN:e3829": "M", "EN:e3830": "N",
        "EN:e3831": "O"
    }

    rekap = defaultdict(lambda: defaultdict(float))

    # Rekap data
    for _, r in df_piutang.iterrows():
        sektor = r["SKTR_KNM_LPNGN_SH"]
        akun   = r["KTGR_SH_KNGN_BRKLNJTN"]
        if akun in MAP_AKUN_COL:
            rekap[sektor][akun] += r["PIUTANG"]

    # Proses baris 3+ - baris 1&2 aman, formula aman
    for row in range(START_ROW, ws.max_row + 1):
        sektor = ws[f"B{row}"].value
        
        for akun, col in MAP_AKUN_COL.items():
            cell = ws[f"{col}{row}"]
            
            # **CEK FORMULA - JANGAN SENTUH!**
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # Formula → Skip
            
            # Data → Isi 0/data
            value = rekap.get(sektor, {}).get(akun, 0)
            cell.value = value

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()
## Golongan Debitur
### 1. Gol. Debitur vs. NPF
def laporan_golongan_vs_klts(template_path, output_path, df_piutang):

    START_ROW = 3

    PIUTANG_COL = {
        "SF:e2": 2,
        "SF:e12": 3,
        "SF:e9": 4,
        "SF:e3": 5,
        "SF:e4": 6
    }

    CKPN_COL = {
        "SF:e2": 9,
        "SF:e12": 10,
        "SF:e9": 11,
        "SF:e3": 12,
        "SF:e4": 13
    }

    FORM_SHEET_MAP = {
        2110: "PS",
        2120: "OK",
        2130: "SUS",
        2140: "PUP"
    }

    # ===============================
    # REKAP DATA
    # ===============================
    rekap_piu  = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    rekap_ckpn = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for _, r in df_piutang.iterrows():
        form = r.get("FORM")
        gol  = r.get("GOLONGAN_DEBITUR")
        klts = r.get("KLTS")

        if form in FORM_SHEET_MAP and pd.notna(gol) and klts in PIUTANG_COL:
            rekap_piu[form][gol][klts] += r.get("PIUTANG", 0)

            # ⛔ PS TIDAK ADA CKPN
            if form != 2110:
                rekap_ckpn[form][gol][klts] += r.get("CKPN", 0)

    # ===============================
    # LOAD TEMPLATE
    # ===============================
    wb = load_workbook(template_path)

    # ===============================
    # ISI PER SHEET
    # ===============================
    for form, sheet_name in FORM_SHEET_MAP.items():

        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        data_piu  = rekap_piu.get(form, {})
        data_ckpn = rekap_ckpn.get(form, {})

        row = START_ROW
        while True:
            golongan = ws.cell(row, 1).value  # kolom A
            if golongan is None:
                break

            # ---- PIUTANG (SEMUA FORM) ----
            for k, c in PIUTANG_COL.items():
                cell = ws.cell(row, c)
                if cell.data_type == "f":
                    continue
                cell.value = data_piu.get(golongan, {}).get(k, 0)

            # ---- CKPN (KECUALI PS) ----
            if form != 2110:
                for k, c in CKPN_COL.items():
                    cell = ws.cell(row, c)
                    if cell.data_type == "f":
                        continue
                    cell.value = data_ckpn.get(golongan, {}).get(k, 0)

            row += 1

    # ===============================
    # SAVE
    # ===============================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 2. Gol. Debitur vs. Jenis Usaha
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook

def laporan_golongan_vs_jenis_usaha(template_path, output_path, df_piutang):

    START_ROW = 3

    # =========================
    # MAP JENIS USAHA → KOLOM
    # =========================
    USAHA_COL = {
        "AK:e20": 3,   # D
        "AK:e21": 4,   # E
        "AK:e22": 5,   # F
        "AK:e23": 6,   # G
        "AK:e24": 8,   # I
        "AK:e25": 9,   # J
        "AK:e26": 10,  # K
        "AK:e27": 11   # L
    }

    # ===============================
    # REKAP DATA
    # ===============================
    rekap = defaultdict(lambda: defaultdict(float))

    for _, r in df_piutang.iterrows():
        gol   = r.get("GOLONGAN_DEBITUR")
        usaha = r.get("JNS_KGTN_SH_PRSHN_MDL_VNTR")
        nilai = r.get("PIUTANG", 0)

        if pd.notna(gol) and usaha in USAHA_COL:
            rekap[gol][usaha] += nilai

    # ===============================
    # LOAD TEMPLATE
    # ===============================
    wb = load_workbook(template_path)
    ws = wb.active   # hanya 1 sheet

    # ===============================
    # ISI KE TEMPLATE
    # ===============================
    row = START_ROW
    while True:
        golongan = ws.cell(row, 1).value  # kolom A
        if golongan is None:
            break

        for usaha, col in USAHA_COL.items():
            cell = ws.cell(row, col)
            if cell.data_type == "f":     # formula aman
                continue
            cell.value = rekap.get(golongan, {}).get(usaha, 0)

        row += 1

    # ===============================
    # SAVE
    # ===============================
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

### 3. Gol. Debitur vs. Identifier
import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def laporan_golongan_vs_identifier(template_path, output_path, df_perusahaan, df_piutang):
    """
    Membuat laporan Identifier vs UMKM dengan sheet Total, NonPerform, dan NPF %.
    
    Args:
        template_path (str): Path template Excel
        output_path (str): Path output Excel
        df_perusahaan_path (str): Path Data Perusahaan.xlsx
        df_piutang_path (str): Path Data Piutang dan CKPN.xlsx
    """
    # Konfigurasi
    START_ROW, START_COL, END_COL = 3, 4, 9
    NUM_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    PERCENT_FORMAT = '0.00%'
    
    FILL_SYARIAH = PatternFill("solid", fgColor="92D050")
    FILL_TOTAL = PatternFill("solid", fgColor="FF99FF")
    BORDER_THIN = Border(left=Side(style="thin"), right=Side(style="thin"), 
                        top=Side(style="thin"), bottom=Side(style="thin"))
    
    GOL_MAP = {"Bank": 4, "LJKNB": 5, "Perusahaan Non Keuangan": 6, "Pemerintah": 7, "Perseorangan": 8}
    
    def apply_border(ws, last_row):
        for r in range(2, last_row + 1):
            for c in range(START_COL, END_COL + 1):
                ws.cell(r, c).border = BORDER_THIN
    
    def fill_base(ws):
        row = START_ROW
        for _, p in df_perusahaan.iterrows():
            ws[f"B{row}"] = p["IDENTIFIER"]
            ws[f"C{row}"] = p["NM_SBTN_PRSHN"]
            
            if p["JNS_PRSHN"] == "Syariah":
                for col in "ABC":
                    ws[f"{col}{row}"].fill = FILL_SYARIAH
            
            row += 1
        return row - 1
    
    def apply_total(ws, last_row, number_format):
        total_row = last_row + 2
        ws[f"C{total_row}"] = "Total"
        ws[f"C{total_row}"].font = Font(bold=True)
        
        for c in range(1, END_COL + 1):
            ws.cell(total_row, c).fill = FILL_TOTAL
        
        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            cell = ws.cell(total_row, c)
            cell.value = f"=SUM({col}{START_ROW}:{col}{last_row})"
            cell.font = Font(bold=True)
            cell.number_format = number_format
        
        return total_row
    
    def fill_value_sheet(ws, df_src):
        rekap = defaultdict(lambda: defaultdict(float))
        for _, r in df_src.iterrows():
            if r["GOLONGAN_DEBITUR"] in GOL_MAP:
                rekap[r["IDENTIFIER"]][r["GOLONGAN_DEBITUR"]] += r["PIUTANG"]
        
        last_row = fill_base(ws)
        
        for r in range(START_ROW, last_row + 1):
            ident = ws[f"B{r}"].value
            for k, c in GOL_MAP.items():
                cell = ws.cell(r, c)
                cell.value = rekap.get(ident, {}).get(k, 0)
                cell.number_format = NUM_FORMAT
            
            ws[f"I{r}"].value = f"=SUM(D{r}:H{r})"
            ws[f"I{r}"].number_format = NUM_FORMAT
        
        total_row = apply_total(ws, last_row, NUM_FORMAT)
        apply_border(ws, last_row)
        return last_row, total_row
    
    # Load template
    wb = load_workbook(template_path)
    
    # Total sheet
    fill_value_sheet(wb["Total"], df_piutang)
    
    # NonPerform sheet
    df_np = df_piutang[df_piutang["KLTS"].isin(["SF:e9", "SF:e3", "SF:e4"])]
    fill_value_sheet(wb["NonPerform"], df_np)
    
    # NPF sheet
    ws_npf = wb["NPF"]
    last_row = fill_base(ws_npf)
    
    for r in range(START_ROW, last_row + 1):
        for c in range(START_COL, END_COL + 1):
            col = get_column_letter(c)
            ws_npf.cell(r, c).value = f"=IFERROR(NonPerform!{col}{r}/Total!{col}{r},0)"
            ws_npf.cell(r, c).number_format = PERCENT_FORMAT
    
    total_row = last_row + 2
    ws_npf[f"C{total_row}"] = "Total"
    ws_npf[f"C{total_row}"].font = Font(bold=True)
    
    for c in range(1, END_COL + 1):
        ws_npf.cell(total_row, c).fill = FILL_TOTAL
    
    for c in range(START_COL, END_COL + 1):
        col = get_column_letter(c)
        cell = ws_npf.cell(total_row, c)
        cell.value = f"=IFERROR(NonPerform!{col}{total_row}/Total!{col}{total_row},0)"
        cell.font = Font(bold=True)
        cell.number_format = PERCENT_FORMAT
    
    apply_border(ws_npf, last_row)
    
    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

# Run Laporan
def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
def run_laporan_npf(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "04. NPF")
    ensure_dir(folder)

    template_path = os.path.join(TEMPLATE_DIR, "Template NPF.xlsx")
    
    output_path = os.path.join(folder, "NPF.xlsx")

    laporan_npf(template_path, output_path, df_perusahaan, df_piutang)
def run_laporan_jenis_pembiayaan(base_dir, df_piutang):
    folder = os.path.join(base_dir, "11. Jenis Pembiayaan")
    os.makedirs(folder, exist_ok=True)

    template_path = os.path.join(TEMPLATE_DIR, "Template Jenis Pembiayaan vs. NPF.xlsx")
   
    output_path = os.path.join(
        folder,
        "Jenis Pembiayaan vs. NPF.xlsx"
    )

    laporan_jenis_pembiayaan_vs_npf(
        template_path,
        output_path,
        df_piutang
    )
def run_laporan_umkm(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "12. UMKM")
    os.makedirs(folder, exist_ok=True)

    laporan_list = [
        {
            "template": "Template UMKM vs. NPF.xlsx",
            "output": "1. UMKM vs. NPF.xlsx",
            "func": laporan_umkm_vs_npf,
            "args": (df_piutang,)
        },
        {
            "template": "Template UMKM vs. Jenis Usaha.xlsx",
            "output": "2. UMKM vs. Jenis Usaha.xlsx",
            "func": laporan_umkm_vs_jenis_usaha,
            "args": (df_piutang,)
        },
        {
            "template": "Template UMKM vs. Jenis Usaha (Kontrak).xlsx",
            "output": "3. UMKM vs. Jenis Usaha (Kontrak).xlsx",
            "func": laporan_umkm_vs_jenis_usaha_kontrak,
            "args": (df_piutang,)
        },
        {
            "template": "Template Identifier vs. UMKM.xlsx",
            "output": "4. Identifier vs. UMKM.xlsx",
            "func": laporan_identifier_vs_umkm,
            "args": (df_perusahaan, df_piutang)
        },
        {
            "template": "Template UMKM vs. Lokasi.xlsx",
            "output": "5. UMKM vs. Lokasi.xlsx",
            "func": laporan_umkm_vs_lokasi,
            "args": (df_piutang,)
        },
        {
            "template": "Template UMKM vs. Lokasi (Kontrak).xlsx",
            "output": "6. UMKM vs. Lokasi (Kontrak).xlsx",
            "func": laporan_umkm_vs_lokasi_kontrak,
            "args": (df_piutang,)
        },
        {
            "template": "Template UMKM vs. Sektor Ekonomi.xlsx",
            "output": "7. UMKM vs. Sektor Ekonomi.xlsx",
            "func": laporan_umkm_vs_sektor_ekonomi,
            "args": (df_piutang,)
        },
    ]

    for lap in laporan_list:
        template_path = os.path.join(TEMPLATE_DIR, lap["template"])
        output_path = os.path.join(folder, lap["output"])

        lap["func"](template_path, output_path, *lap["args"])

def run_laporan_pinjaman_valas(base_dir, df_perusahaan, df_pinjaman_valas):
    folder = os.path.join(base_dir, "13. Pinjaman Valas")
    os.makedirs(folder, exist_ok=True)

    template_path = os.path.join(TEMPLATE_DIR, "Template Pinjaman Valas.xlsx")

    output_path = os.path.join(folder, "Pinjaman Valas.xlsx")

    laporan_pinjaman_valas(
        template_path,
        output_path,
        df_perusahaan,
        df_pinjaman_valas
    )

def run_laporan_piutang_valas(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "14. Piutang Valas")
    os.makedirs(folder, exist_ok=True)

    template_path = os.path.join(TEMPLATE_DIR, "Template Piutang Valas.xlsx")

    output_path = os.path.join(folder, "Piutang Valas.xlsx")

    laporan_piutang_valas(
        template_path,
        output_path,
        df_perusahaan,
        df_piutang
    )
def run_laporan_golongan_debitur(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "15. Golongan Debitur")
    os.makedirs(folder, exist_ok=True)

    laporan_list = [
        {
            "template": "Template Gol. Debitur vs. NPF.xlsx",
            "output": "1. Gol. Debitur vs. NPF.xlsx",
            "func": laporan_golongan_vs_klts,
            "args": (df_piutang,)
        },
        {
            "template": "Template Gol. Debitur vs. Jenis Usaha.xlsx",
            "output": "2. Gol. Debitur vs. Jenis Usaha.xlsx",
            "func": laporan_golongan_vs_jenis_usaha,
            "args": (df_piutang,)
        },
        {
            "template": "Template Gol. Debitur vs. Identifier.xlsx",
            "output": "3. Gol. Debitur vs. Identifier.xlsx",
            "func": laporan_golongan_vs_identifier,
            "args": (df_perusahaan, df_piutang)
        },
    ]

    for lap in laporan_list:
        template_path = os.path.join(TEMPLATE_DIR, lap["template"])
        output_path = os.path.join(folder, lap["output"])

        # validasi template (penting untuk GUI)
        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"Template tidak ditemukan: {template_path}"
            )

        lap["func"](template_path, output_path, *lap["args"])

def run_laporan_jumlah_kontrak(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "16. Jumlah Kontrak")
    os.makedirs(folder, exist_ok=True)

    laporan_list = [
        {
            "template": "Template Identifier vs. Jenis Usaha.xlsx",
            "output": "1. Identifier vs. Jenis Usaha.xlsx",
            "func": laporan_identifier_vs_jenis_usaha_kontrak,
            "args": (df_perusahaan, df_piutang)
        },
        {
            "template": "Template Identifier vs. Kolektibilitas.xlsx",
            "output": "2. Identifier vs. Kolektibilitas.xlsx",
            "func": laporan_identifier_vs_kolektibilitas,
            "args": (df_perusahaan, df_piutang)
        },
        {
            "template": "Template Jenis Usaha vs. Kolektibilitas.xlsx",
            "output": "3. Jenis Usaha vs. Kolektibilitas.xlsx",
            "func": laporan_jenis_pembiayaan_vs_kolektibilitas,
            "args": (df_piutang,)
        },
        {
            "template": "Template Lokasi vs. Jenis Usaha.xlsx",
            "output": "4. Lokasi vs. Jenis Usaha (Kontrak).xlsx",
            "func": laporan_lokasi_vs_jenis_usaha_kontrak,
            "args": (df_piutang,)
        },
        {
            "template": "Template Sektor Ekonomi vs. Jenis Usaha (Kontrak).xlsx",
            "output": "5. Sektor Ekonomi vs. Jenis Usaha (Kontrak).xlsx",
            "func": laporan_sektor_vs_jenis_usaha_kontrak,
            "args": (df_piutang,)
        },
    ]

    for lap in laporan_list:
        template_path = os.path.join(TEMPLATE_DIR, lap["template"])
        output_path = os.path.join(folder, lap["output"])

        # Validasi template (penting untuk GUI & user non-teknis)
        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"Template tidak ditemukan: {template_path}"
            )

        lap["func"](template_path, output_path, *lap["args"])

def run_laporan_lokasi(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "17. Lokasi")
    os.makedirs(folder, exist_ok=True)

    laporan_list = [
        {
            "template": "Template Lokasi vs. NPF.xlsx",
            "output": "1. Lokasi vs. NPF.xlsx",
            "func": laporan_lokasi_vs_npf,
            "args": (df_piutang,)
        },
        {
            "template": "Template Lokasi vs. Jenis Usaha (NPF).xlsx",
            "output": "2. Lokasi vs. Jenis Usaha (NPF).xlsx",
            "func": laporan_lokasi_vs_jenis_usaha_npf,
            "args": (df_piutang,)
        },
        {
            "template": "Template Lokasi vs. Identifier.xlsx",
            "output": "3. Lokasi vs. Identifier.xlsx",
            "func": laporan_lokasi_vs_identifier,
            "args": (df_perusahaan, df_piutang)
        },
        {
            "template": "Template Lokasi vs. Identifier.xlsx",
            "output": "4. Lokasi vs. Identifier (Kontrak).xlsx",
            "func": laporan_lokasi_vs_identifier_kontrak,
            "args": (df_perusahaan, df_piutang)
        },
        {
            "template": "Template Lokasi vs. Sektor Ekonomi.xlsx",
            "output": "5. Lokasi vs. Sektor Ekonomi.xlsx",
            "func": laporan_lokasi_vs_sektor,
            "args": (df_piutang,)
        },
    ]

    for lap in laporan_list:
        template_path = os.path.join(TEMPLATE_DIR, lap["template"])
        output_path = os.path.join(folder, lap["output"])

        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"Template tidak ditemukan: {template_path}"
            )

        lap["func"](template_path, output_path, *lap["args"])

def run_laporan_sektor_ekonomi(base_dir, df_perusahaan, df_piutang):
    folder = os.path.join(base_dir, "18. Sektor Ekonomi")
    ensure_dir(folder)

    laporan_list = [
        {
            "template": "Template Sektor Ekonomi vs. NPF.xlsx",
            "output": "1. Sektor Ekonomi vs. NPF.xlsx",
            "func": laporan_sektor_vs_npf,
            "args": (df_piutang,)
        },
        {
            "template": "Template Sektor Ekonomi vs. Jenis Usaha.xlsx",
            "output": "2. Sektor Ekonomi vs. Jenis Usaha.xlsx",
            "func": laporan_sektor_vs_jenis_usaha,
            "args": (df_piutang,)
        },
        {
            "template": "Template Sektor Ekonomi vs. Identifier.xlsx",
            "output": "3. Sektor Ekonomi vs. Identifier.xlsx",
            "func": laporan_sektor_vs_identifier,
            "args": (df_perusahaan, df_piutang)
        },
        {
            "template": "Template Identifier vs. Sektor Ekonomi.xlsx",
            "output": "4. Identifier vs. Sektor Ekonomi (NPF).xlsx",
            "func": laporan_identifier_vs_sektor,
            "args": (df_perusahaan, df_piutang)
        },
    ]

    for lap in laporan_list:
        template_path = os.path.join(TEMPLATE_DIR, lap["template"])
        output_path = os.path.join(folder, lap["output"])

        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"Template tidak ditemukan: {template_path}"
            )

        lap["func"](template_path, output_path, *lap["args"])

def run_laporan_kukb(base_dir, df_piutang):
    folder = os.path.join(
        base_dir,
        "19. Kategori Usaha Keuangan Berkelanjutan"
    )
    os.makedirs(folder, exist_ok=True)

    laporan_list = [
        {
            "template": "Template KUKB vs. NPF.xlsx",
            "output": "1. KUKB vs. NPF.xlsx",
            "func": laporan_kukb_vs_npf,
            "args": (df_piutang,)
        },
        {
            "template": "Template KUKB vs. Pariwisata.xlsx",
            "output": "2. KUKB vs. Pariwisata.xlsx",
            "func": laporan_kukb_vs_pariwisata,
            "args": (df_piutang,)
        },
        {
            "template": "Template KUKB vs. Ekonomi Kreatif.xlsx",
            "output": "3. KUKB vs. Ekonomi Kreatif.xlsx",
            "func": laporan_kukb_vs_ekraf,
            "args": (df_piutang,)
        },
    ]

    for lap in laporan_list:
        template_path = os.path.join(TEMPLATE_DIR, lap["template"])
        output_path = os.path.join(folder, lap["output"])

        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"Template tidak ditemukan: {template_path}"
            )

        lap["func"](template_path, output_path, *lap["args"])

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def generate_output_path(tipe_laporan, laporan_dir, period):
    cfg = LAPORAN_CONFIG[tipe_laporan.lower()]
    folder_name = cfg["output_folder"]

    output_folder = os.path.join(laporan_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)

    laporan_name = folder_name.split(". ", 1)[1]
    file_name = f"{laporan_name} {period}.xlsx"

    return os.path.join(output_folder, file_name)

def find_excel_file(base_path, form_code):
    """Cari file Excel berdasarkan 4 angka pertama nama file = form_code"""
    pattern = os.path.join(base_path, "*.xlsx")
    files = glob.glob(pattern)
    
    for file in files:
        filename = os.path.basename(file)
        # Ambil 4 angka pertama dari nama file
        match = re.match(r'^(\d{4})', filename)
        if match and match.group(1) == form_code:
            return file
        
def proses_laporan_config(tipe_laporan, data_dir, laporan_dir, period):
    tipe_laporan = tipe_laporan.lower()

    if tipe_laporan not in LAPORAN_CONFIG:
        raise ValueError(f"Tipe laporan '{tipe_laporan}' tidak tersedia")

    cfg = LAPORAN_CONFIG[tipe_laporan]

    form_code   = cfg["form"]
    template    = cfg["template"]
    sheet_name = cfg["sheet"]
    transform  = cfg["transform"]

    # ==================================================
    # 1. Cari file sumber dari data_dir
    # ==================================================
    excel_file = find_excel_file(data_dir, form_code)
    print(f"📁 Sumber data: {excel_file}")

    # ==================================================
    # 2. Baca sheet utama
    # ==================================================
    df = pd.read_excel(excel_file, sheet_name=form_code, engine="openpyxl")

    # ==================================================
    # 3. Baca sheet sya (opsional)
    # ==================================================
    try:
        df_sya = pd.read_excel(excel_file, sheet_name="sya", engine="openpyxl")
        if transform:
            df_sya = transform(df_sya)
    except:
        df_sya = pd.DataFrame()

    # ==================================================
    # 4. Gabungkan data
    # ==================================================
    if not df_sya.empty:
        df_sya = df_sya.reindex(columns=df.columns, fill_value=0)
        df_gab = pd.concat([df, df_sya], ignore_index=True)
    else:
        df_gab = df.copy()

    # ==================================================
    # 5. Bersihkan & agregasi
    # ==================================================
    ignore_cols = ["ID", "PERIOD", "IDENTIFIER"]
    var_cols = [c for c in df_gab.columns if c not in ignore_cols]

    df_gab[var_cols] = (
        df_gab[var_cols]
        .apply(pd.to_numeric, errors="coerce")
        .replace([np.inf, -np.inf], np.nan)
        .fillna(0)
    )

    value_map = df_gab[var_cols].sum().to_dict()

    # ==================================================
    # 6. Isi template
    # ==================================================
    wb = load_workbook(template, data_only=False)
    ws = wb[sheet_name]

    updated = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in value_map:
                cell.value = value_map[cell.value]
                updated += 1

    # ==================================================
    # 7. Simpan output
    # ==================================================
    output_path = generate_output_path(tipe_laporan, laporan_dir, period)
    wb.save(output_path)
    wb.close()

    print(f"✅ {tipe_laporan.upper()} selesai")
    print(f"📄 Output : {output_path}")
    print(f"🧮 Cell di-update : {updated}")

def run_all_ref_laporan(base_dir, df_perusahaan, df_piutang, df_pinjaman_valas):
    """Jalankan semua laporan dari data referensi"""
    
    # Urutan eksekusi sesuai folder number
    run_functions = [
        lambda: run_laporan_npf(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_jenis_pembiayaan(base_dir, df_piutang),
        lambda: run_laporan_umkm(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_pinjaman_valas(base_dir, df_perusahaan, df_pinjaman_valas),
        lambda: run_laporan_piutang_valas(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_golongan_debitur(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_jumlah_kontrak(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_lokasi(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_sektor_ekonomi(base_dir, df_perusahaan, df_piutang),
        lambda: run_laporan_kukb(base_dir, df_piutang)
    ]
    
    for run_func in run_functions:
        try:
            run_func()
        except Exception as e:
            print(f"Warning: Laporan gagal {e}")
            continue


import os
import pandas as pd

# ========================================
# FORM RULES
# ========================================
FORM_RULES = {
    "2110": {"piutang": "NL_PNYRTN_SHM_PRD_LPRN", "ckpn": None, "sya_override": "AK:e24", "kontrak": 1},
    "2120": {"piutang": "SLD_KHR_PRD_PLPRN", "ckpn": ["NL_ST_BK_CDNGN_KRGN_PNRNN_NL","NL_ST_KRNG_BK_CDNGN_KRGN_PNRNN_NL","NL_ST_TDK_BK_CDNGN_KRGN_PNRNN_NL"], "sya_override": "AK:e25", "kontrak": 1},
    "2130": {"piutang": "SLD_KHR_PRD_PLPRN", "ckpn": ["NL_ST_BK_CDNGN_KRGN_PNRNN_NL","NL_ST_KRNG_BK_CDNGN_KRGN_PNRNN_NL","NL_ST_TDK_BK_CDNGN_KRGN_PNRNN_NL"], "sya_override": "AK:e26", "kontrak": 1},
    "2140": {"piutang": "PTNG_PMBYN__PKK", "ckpn": ["NL_ST_BK_CDNGN_KRGN_PNRNN_NL","NL_ST_KRNG_BK_CDNGN_KRGN_PNRNN_NL","NL_ST_TDK_BK_CDNGN_KRGN_PNRNN_NL"], "sya_override": None, "kontrak": "KONTRAK"},
    "2550": {"special": True}  # Special handling - no piutang/ckpn
}


def extract_form_code(filename):
    """Extract 4 digit TERDEPAN dari nama file"""
    if len(filename) >= 4 and filename[:4].isdigit():
        return filename[:4]
    return None


def normalize_df(df, form, sheet_name, piutang_col, ckpn_cols=None, jns_kgt_override=None, kontrak_rule=None):
    """Transformasi sesuai FORM_RULES - TANPA filter header, TANPA kolom FORM"""
    df = df.copy()
    
    # PIUTANG
    df["PIUTANG"] = df[piutang_col].fillna(0) if piutang_col and piutang_col in df.columns else 0
    df["FORM"] = form
    # CKPN
    if ckpn_cols:
        cols_exist = [c for c in ckpn_cols if c in df.columns]
        df["CKPN"] = df[cols_exist].sum(axis=1).fillna(0) if cols_exist else 0
    else:
        df["CKPN"] = 0
    
    # KONTRAK
    if isinstance(kontrak_rule, str) and kontrak_rule in df.columns:
        df["KONTRAK"] = df[kontrak_rule].fillna(0)
    else:
        df["KONTRAK"] = kontrak_rule
    
    # SYA override
    if sheet_name == "sya" and jns_kgt_override:
        df["JNS_KGTN_SH_PRSHN_MDL_VNTR"] = jns_kgt_override
    
    # 12 kolom standar TANPA FORM
    cols_final = [
        "JNS_KGTN_SH_PRSHN_MDL_VNTR", "GLNGN_PHK_LWN", "SKTR_KNM_LPNGN_SH",
        "KTGR_SH_KNGN_BRKLNJTN", "KTGR_SH_DBTR", "LKS_DT__PRYK",
        "KLTS", "JNS_MT_NG", "KONTRAK", "PIUTANG", "CKPN", "IDENTIFIER", "FORM"
    ]
    
    return df.reindex(columns=cols_final)


def normalize_df_2550(df, form, sheet_name):
    """Special normalization untuk FORM 2550 - hanya 3 kolom spesifik"""
    df = df.copy()
    
    # Pastikan kolom 2550 ada
    required_cols = ["JNS_MT_NG", "SLD_PNJMN", "IDENTIFIER"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        print(f"   ⚠️  Missing columns for 2550: {missing_cols}")
        return pd.DataFrame()
    
    # Buat DataFrame minimal TANPA FORM
    df_norm = pd.DataFrame({
        "JNS_MT_NG": df["JNS_MT_NG"],
        "SLD_PNJMN": df["SLD_PNJMN"],
        "IDENTIFIER": df["IDENTIFIER"]
    })
    
    # Fill NA dengan 0 untuk SLD_PNJMN
    df_norm["SLD_PNJMN"] = df_norm["SLD_PNJMN"].fillna(0)
    
    return df_norm


def process_folder(folder_path, output_dir):
    """Save ke output_dir (script folder)"""
    piutang_ckpn_dfs = []
    pinjaman_valas_dfs = []
    
    # ✅ SEMUA ke script_dir
    piutang_ckpn_path = os.path.join(output_dir, "Data Piutang dan CKPN.xlsx")
    pinjaman_valas_path = os.path.join(output_dir, "Data Pinjaman Valas.xlsx")
    
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"🔍 Scanning: {folder_path}")
    print("-" * 60)
    
    # Target FORM codes
    target_forms = ["2110", "2120", "2130", "2140", "2550"]
    valid_files = []
    
    # SCAN FILES
    for fname in os.listdir(folder_path):
        if fname.lower().endswith(".xlsx"):
            form = extract_form_code(fname)
            if form in target_forms:
                valid_files.append((fname, form))
                print(f"✅ Found {form}: {fname}")
    
    if not valid_files:
        print("❌ No 2110-2140/2550 files found!")
        return pd.DataFrame()
    
    print(f"\n🚀 Processing {len(valid_files)} valid files...")
    print("-" * 60)
    
    # PROCESS VALID FILES
    for fname, form in valid_files:
        rule = FORM_RULES.get(form, {})
        fpath = os.path.join(folder_path, fname)
        
        print(f"\n🔄 [{form}] {fname}")
        
        try:
            xls = pd.ExcelFile(fpath)
            print(f"   📋 Sheets: {xls.sheet_names}")
        except Exception as e:
            print(f"   ❌ Error: {e}")
            continue
        
        # Process logic berdasarkan FORM type
        if form == "2550":
            # SPECIAL 2550 -> Data_Pinjaman_Valas
            target_sheets = ["2550", "sya"]
            for sheet in target_sheets:
                if sheet not in xls.sheet_names:
                    continue
                
                df_raw = pd.read_excel(xls, sheet_name=sheet)
                print(f"   📄 {sheet}: {len(df_raw)} rows loaded")
                
                df_norm = normalize_df_2550(df_raw, form, sheet)
                df_clean = df_norm.dropna(how='all')
                
                if len(df_clean) > 0:
                    pinjaman_valas_dfs.append(df_clean)
                    print(f"   ✅ {sheet}: {len(df_clean)} rows added -> Pinjaman Valas")
        else:
            # STANDARD 2110-2140 -> Data_Piutang_dan_CKPN
            target_sheets = [form, "sya"]
            for sheet in target_sheets:
                if sheet not in xls.sheet_names:
                    continue
                
                df_raw = pd.read_excel(xls, sheet_name=sheet)
                print(f"   📄 {sheet}: {len(df_raw)} rows loaded")
                
                df_norm = normalize_df(
                    df=df_raw,
                    form=form,
                    sheet_name=sheet,
                    piutang_col=rule.get("piutang"),
                    ckpn_cols=rule.get("ckpn"),
                    jns_kgt_override=rule.get("sya_override"),
                    kontrak_rule=rule.get("kontrak")
                )
                
                df_clean = df_norm.dropna(how='all')
                if len(df_clean) > 0:
                    piutang_ckpn_dfs.append(df_clean)
                    print(f"   ✅ {sheet}: {len(df_clean)} rows added -> Piutang & CKPN")
    
    # SAVE FILES TERPISAH
    print("\n" + "="*60)
    
    # 1. Data Piutang dan CKPN (2110-2140)
    if piutang_ckpn_dfs:
        piutang_ckpn_df = pd.concat(piutang_ckpn_dfs, ignore_index=True)
        piutang_ckpn_df.to_excel(piutang_ckpn_path, index=False, engine='openpyxl')
        print(f"🎉 Piutang & CKPN SUCCESS!")
        print(f"📊 Total: {len(piutang_ckpn_df)} rows")
        print(f"💾 Saved: {piutang_ckpn_path}")
        print(f"📋 Columns: {list(piutang_ckpn_df.columns)}")   
    else:
        print("❌ No Piutang & CKPN data!")
    
    # 2. Data Pinjaman Valas (2550)
    if pinjaman_valas_dfs:
        pinjaman_valas_df = pd.concat(pinjaman_valas_dfs, ignore_index=True)
        pinjaman_valas_df.to_excel(pinjaman_valas_path, index=False,  engine='openpyxl')
        print(f"🎉 Pinjaman Valas SUCCESS!")
        print(f"📊 Total: {len(pinjaman_valas_df)} rows")
        print(f"💾 Saved: {pinjaman_valas_path}")
        print(f"📋 Columns: {list(pinjaman_valas_df.columns)}")
    else:
        print("❌ No Pinjaman Valas data!")
    
    return {
        "piutang_ckpn": piutang_ckpn_df if 'piutang_ckpn_df' in locals() else pd.DataFrame(),
        "pinjaman_valas": pinjaman_valas_df if 'pinjaman_valas_df' in locals() else pd.DataFrame()
    }

def preprocess_data(data_dir, ref_dir, template_dir):
    """
    Update: 
    1. Extract Piutang/CKPN & Pinjaman Valas dari data_dir (2110/2120/2130/2140/2550)
    2. Load Data Perusahaan dari data_ref_dir (existing)
    3. Apply mapping sektor & golongan (existing)
    4. Save intermediate files ke data_ref_dir
    """
    print("📊 Preprocessing data dari data_dir...")
    
    # =========================
    # 1. EXTRACT dari data_dir menggunakan process_folder logic
    # =========================
    extract_result = process_folder(data_dir, output_dir=ref_dir)
    df_piutang_ckpn = extract_result["piutang_ckpn"]
    df_pinjaman_valas = extract_result["pinjaman_valas"]
    
    if df_piutang_ckpn.empty:
        raise ValueError("❌ Tidak ada data Piutang/CKPN dari form 2110-2140 di data_dir!")
    if df_pinjaman_valas.empty:
        print("⚠️  Tidak ada data Pinjaman Valas (2550) - lanjut tanpa valas")
    
    print(f"✅ Extracted: {len(df_piutang_ckpn)} rows Piutang/CKPN, {len(df_pinjaman_valas)} rows Valas")
    
    # =========================
    # 2. LOAD Data Perusahaan dari data_ref_dir (EXISTING)
    # =========================
    df_perusahaan_path = os.path.join(ref_dir, "Data Perusahaan.xlsx")
    if not os.path.exists(df_perusahaan_path):
        raise ValueError(f"❌ Data Perusahaan.xlsx tidak ditemukan di {ref_dir}")
    
    df_perusahaan = pd.read_excel(df_perusahaan_path)
    print(f"✅ Loaded Data Perusahaan: {len(df_perusahaan)} rows")
    
    # =========================
    # 3. TRANSFORM PIUTANG (EXISTING LOGIC)
    # =========================
    df_piutang = df_piutang_ckpn.copy()
    
    # --- Mapping sektor ekonomi ---
    sektor_map = {
        "SE:e01":"SE:A","SE:e02":"SE:A","SE:e03":"SE:A",
        "SE:e05":"SE:B","SE:e06":"SE:B","SE:e07":"SE:B","SE:e08":"SE:B","SE:e09":"SE:B",
        "SE:e10":"SE:C","SE:e11":"SE:C","SE:e12":"SE:C","SE:e13":"SE:C","SE:e14":"SE:C",
        "SE:e15":"SE:C","SE:e16":"SE:C","SE:e17":"SE:C","SE:e18":"SE:C","SE:e19":"SE:C",
        "SE:e20":"SE:C","SE:e21":"SE:C","SE:e22":"SE:C","SE:e23":"SE:C","SE:e24":"SE:C",
        "SE:e25":"SE:C","SE:e26":"SE:C","SE:e27":"SE:C","SE:e28":"SE:C","SE:e29":"SE:C",
        "SE:e30":"SE:C","SE:e31":"SE:C","SE:e32":"SE:C","SE:e33":"SE:C",
        "SE:e35":"SE:D",
        "SE:e36":"SE:E","SE:e37":"SE:E","SE:e38":"SE:E","SE:e39":"SE:E",
        "SE:e41":"SE:F","SE:e42":"SE:F","SE:e43":"SE:F",
        "SE:e45":"SE:G","SE:e46":"SE:G","SE:e47":"SE:G",
        "SE:e49":"SE:H","SE:e50":"SE:H","SE:e51":"SE:H","SE:e52":"SE:H","SE:e53":"SE:H",
        "SE:e55":"SE:I","SE:e56":"SE:I",
        "SE:e58":"SE:J","SE:e59":"SE:J","SE:e60":"SE:J","SE:e61":"SE:J","SE:e62":"SE:J","SE:e63":"SE:J",
        "SE:e64":"SE:K","SE:e65":"SE:K","SE:e66":"SE:K",
        "SE:e68":"SE:L",
        "SE:e69":"SE:M","SE:e70":"SE:M","SE:e71":"SE:M","SE:e72":"SE:M","SE:e73":"SE:M","SE:e74":"SE:M","SE:e75":"SE:M",
        "SE:e77":"SE:N","SE:e78":"SE:N","SE:e79":"SE:N","SE:e80":"SE:N","SE:e81":"SE:N","SE:e82":"SE:N",
        "SE:e84":"SE:O","SE:e85":"SE:P",
        "SE:e86":"SE:Q","SE:e87":"SE:Q","SE:e88":"SE:Q",
        "SE:e90":"SE:R","SE:e91":"SE:R","SE:e92":"SE:R","SE:e93":"SE:R",
        "SE:e94":"SE:S","SE:e95":"SE:S","SE:e96":"SE:S",
        "SE:e97":"SE:T","SE:e98":"SE:T",
        "SE:e99":"SE:U",
        "SE:e001":"SE:V","SE:e002":"SE:V","SE:e003":"SE:V","SE:e004":"SE:V",
        "SE:e009":"SE:W"
    }
    df_piutang["SEKTOR_KAT"] = df_piutang["SKTR_KNM_LPNGN_SH"].map(
        lambda x: sektor_map.get(
            x[:7] if str(x).startswith("SE:e00") else x[:6]
        )
    )
    
    # --- Mapping golongan debitur ---
    df_gol = pd.read_excel(
        os.path.join(template_dir, "Jenis Golongan.xlsx"),
        usecols=["Sandi", "Jenis Golongan"]
    )
    
    gol_map = dict(zip(df_gol["Sandi"], df_gol["Jenis Golongan"]))
    df_piutang["GOLONGAN_DEBITUR"] = df_piutang["GLNGN_PHK_LWN"].map(gol_map)
    
    # =========================
    # 4. RETURN seperti sebelumnya
    # =========================
    return {
        "piutang": df_piutang,
        "perusahaan": df_perusahaan,
        "valas": df_pinjaman_valas
    }

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# =============================
# DAFTAR LAPORAN
# =============================
LAPORAN_LIST = [
    # ================================
    # LAPORAN UTAMA (Existing)
    # ================================
    ("Neraca", "neraca"),
    ("Laba Rugi", "laba_rugi"),
    ("Arus Kas", "arus_kas"),
    ("Rekening Administratif", "rekening_administratif"),
    ("Tingkat Pendidikan", "tingkat_pendidikan"),
    ("Tingkat Divisi", "tingkat_divisi"),
    ("LKAL", "laporan_kesesuaian_aset_dan_liabilitas"),
   
    # ================================
    # LAPORAN DATA REFERENSI (BARU)
    # ================================
    ("NPF", "npf"),
    ("Jenis Pembiayaan", "jenis_pembiayaan"),
    ("UMKM", "umkm"),
    ("Pinjaman Valas", "pinjaman_valas"),
    ("Piutang Valas", "piutang_valas"),
    ("Golongan Debitur", "golongan_debitur"),
    ("Jumlah Kontrak", "jumlah_kontrak"),
    ("Lokasi", "lokasi"),
    ("Sektor Ekonomi", "sektor_ekonomi"),
    ("KUKB", "kukb")
]
from tkinter import scrolledtext, filedialog
import tkinter as tk
from tkinter import ttk
import os
from datetime import datetime
import glob

# =============================
# GUI APP - FIXED LENGKAP
# =============================
class LaporanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generator Laporan PMV")
        self.geometry("715x700")  # BESAR untuk log
        self.resizable(True, True)  # BISA RESIZE
        
        self.create_widgets()

    def create_widgets(self):
        # =============================
        # FRAME INPUT
        # =============================
        frame_input = ttk.LabelFrame(self, text="Parameter")
        frame_input.pack(fill="x", padx=10, pady=10)

        # Folder Data
        ttk.Label(frame_input, text="Folder Sumber Data").grid(row=0, column=0, sticky="w")
        self.data_dir_var = tk.StringVar()
        ttk.Entry(frame_input, textvariable=self.data_dir_var, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(frame_input, text="Browse", command=self.browse_data).grid(row=0, column=2)

        # Folder Output
        ttk.Label(frame_input, text="Folder Output").grid(row=1, column=0, sticky="w")
        self.output_dir_var = tk.StringVar()
        ttk.Entry(frame_input, textvariable=self.output_dir_var, width=70).grid(row=1, column=1, padx=5)
        ttk.Button(frame_input, text="Browse", command=self.browse_output).grid(row=1, column=2)

        # Folder Data Referensi
        ttk.Label(frame_input, text="Folder Data Ref.").grid(row=2, column=0, sticky="w", pady=2)
        self.ref_dir_var = tk.StringVar()
        ttk.Entry(frame_input, textvariable=self.ref_dir_var, width=70).grid(row=2, column=1, padx=5, pady=2)
        ttk.Button(frame_input, text="Browse", command=self.browse_ref).grid(row=2, column=2, pady=2)

        # Period
        ttk.Label(frame_input, text="Nama File Laporan").grid(row=4, column=0, sticky="w")
        self.period_var = tk.StringVar()
        ttk.Entry(frame_input, textvariable=self.period_var, width=40).grid(row=4, column=1, sticky="w", padx=5)

        # =============================
        # FRAME LAPORAN
        # =============================
        frame_laporan = ttk.LabelFrame(self, text="Jenis Laporan")
        frame_laporan.pack(fill="x", padx=10, pady=5)

        self.laporan_vars = {}
        for i, (label, key) in enumerate(LAPORAN_LIST):
            var = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(frame_laporan, text=label, variable=var)
            chk.grid(row=i // 2, column=i % 2, sticky="w", padx=10, pady=2)
            self.laporan_vars[key] = var

        # =============================
        # BUTTON
        # =============================
        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(button_frame, text="▶ Proses Laporan", command=self.run_process).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="🗑️ Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)

        # =============================
        # LOG SECTION - SCROLLABLE BESAR
        # =============================
        self.create_log_section()

    def create_log_section(self):
        """Log scrollable BESAR dengan clear button"""
        self.log_frame = ttk.LabelFrame(self, text="📋 Log Proses Laporan", padding=5)
        self.log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
        
        # ScrolledText BESAR
        self.log_text = scrolledtext.ScrolledText(
            self.log_frame,
            height=25,
            width=100,
            font=("Consolas", 10),
            wrap=tk.WORD,
            bg="#f8f9fa",
            fg="#2c3e50",
            state='normal'
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def log(self, message):
        """Log dengan auto-scroll"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # Auto-scroll ke bawah
        self.log_text.config(state='disabled')
        self.log_text.update()

    def clear_log(self):
        """Clear semua log"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.log("🗑️ Log di-clear - Siap proses baru!")

    # =============================
    # HANDLER
    # =============================
    def browse_data(self):
        folder = filedialog.askdirectory()
        if folder:
            self.data_dir_var.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir_var.set(folder)

    def browse_ref(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ref_dir_var.set(folder)

    def run_process(self):
        from datetime import datetime
        import glob
        
        self.log("\n🚀 MULAI GENERATE KK PMV")
        self.log("⏱️  Mulai: " + datetime.now().strftime("%H:%M:%S"))
        self.log("=" * 65)
        
        # INPUT & FOLDER
        data_dir = self.data_dir_var.get()
        ref_dir = self.ref_dir_var.get()
        base_dir = self.output_dir_var.get()
        period = self.period_var.get()
        
        if not all([data_dir, ref_dir, base_dir, period]):
            self.log("❌ SEMUA FIELD WAJIB DIISI!")
            return
        
        script_dir = os.path.join(APP_DIR, "data")
        laporan_dir = os.path.join(base_dir, f"{period}")
        
        # BUAT FOLDER
        os.makedirs(script_dir, exist_ok=True)
        os.makedirs(laporan_dir, exist_ok=True)
        
        # STEP 1: VALIDASI DATA SOURCE
        self.log("📂 1. Validasi data PMV...")
        if not os.path.exists(data_dir):
            self.log(f"   ❌ Folder hilang: {os.path.basename(data_dir)}")
            return
        
        self.log(f"   📁 Folder OK: {os.path.basename(data_dir)}")
        
        # CEK FORM + PIUTANG TERPISAH
        forms = ["1100", "1200", "1300", "1110", "5310", "0043", "0041"]
        piutang_forms = ["2110", "2120", "2130", "2140"]
        
        self.log("   📋 File PMV ditemukan:")
        total_files = 0
        
        for form in forms:
            files = glob.glob(os.path.join(data_dir, f"{form}*.xlsx"))
            total_files += len(files)
            self.log(f"      {form}: {'✓' if files else '✗'} ({len(files)} file)")
        
        self.log("   📋 Piutang dan CKPN (source Tarikan Data {period}):")
        for form in piutang_forms:
            files = glob.glob(os.path.join(data_dir, f"{form}*.xlsx"))
            total_files += len(files)
            self.log(f"      {form}: {'✓' if files else '✗'} ({len(files)} file)")
        
        valas_files = glob.glob(os.path.join(data_dir, "2550*.xlsx"))
        total_files += len(valas_files)
        self.log(f"      2550: {'✓' if valas_files else '✗'} ({len(valas_files)} file)")
        
        self.log(f"   📊 Total: {total_files} file PMV ✓")
        
        # STEP 2: OLAH DATA
        self.log("\n📊 2. Olah Data Piutang dan CKPN & Pinjaman Valas...")
        try:
            data_ref = preprocess_data(data_dir, script_dir, TEMPLATE_DIR)
            self.log(f"   ✅ Data Piutang dan CKPN: {len(data_ref['piutang']):,} baris ✓")
            self.log(f"   ✅ Data Perusahaan: {len(data_ref['perusahaan']):,} baris ✓")
            self.log(f"   ✅ Data Valas: {len(data_ref['valas']):,} baris ✓")
        except Exception as e:
            self.log(f"   ❌ Olah data gagal: {str(e)}")
            return
        
        # STEP 3: SEMUA LAPORAN
        self.log("\n📋 3. Generate KK PMV...")
        selesai = 0
        
        for laporan_name, config in LAPORAN_CONFIG.items():
            if self.laporan_vars[laporan_name].get():
                form_code = config["form"]
                self.log(f"   📄 {laporan_name.replace('_',' ').title()} (form {form_code})...")
                try:
                    proses_laporan_config(laporan_name, data_dir, laporan_dir, period)
                    selesai += 1
                    self.log(f"      ✅ {config['output_folder']} OK")
                except Exception as e:
                    self.log(f"      ❌ {laporan_name}: {str(e)}")
        
        self.log("   📊 Laporan NPF s.d. KUKB ...")
        if not data_ref['piutang'].empty and not data_ref['perusahaan'].empty:
            try:
                run_all_ref_laporan(laporan_dir, data_ref["perusahaan"], data_ref["piutang"], data_ref["valas"])
                selesai += 1
                self.log("      ✅ Laporan NPF s.d. KUKB OK")
            except Exception as e:
                self.log(f"      ❌ Laporan NPF s.d. KUKB: {str(e)}")
        
        # SUMMARY
        self.log("\n" + "="*65)
        self.log("🎉 KK PMV LENGKAP SELESAI!")
        self.log(f"📁 Folder: {laporan_dir}")
        self.log(f"📄 Total laporan: {selesai} file")
        self.log("⏱️  Selesai: " + datetime.now().strftime("%H:%M:%S"))

# =============================
# RUN
# =============================
if __name__ == "__main__":
    app = LaporanApp()
    app.mainloop()
