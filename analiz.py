import pandas as pd
import os
import glob
import re


def analiz_yap():
    # 1. Klasördeki ilk .tsv dosyasını bul
    tsv_dosyalari = glob.glob("*.tsv")

    if not tsv_dosyalari:
        print("Hata: Klasörde .tsv uzantılı dosya bulunamadı!")
        return

    dosya_yolu = tsv_dosyalari[0]
    dosya_adi = os.path.basename(dosya_yolu)

    # Dosya ismindeki sayısal kodları çek (Örn: GSE12345.tsv -> 12345)
    sayisal_kod = "".join(re.findall(r'\d+', dosya_adi))
    ek_isim = f"_{sayisal_kod}" if sayisal_kod else ""

    print(f"{dosya_adi} dosyası işleniyor... (Kod: {sayisal_kod})")

    # 2. Dosyayı oku
    df = pd.read_csv(dosya_yolu, sep='\t')

    # 3. Sütun isimlerini eşleştirme
    padj_col = next((c for c in df.columns if c.lower() in ['padj', 'adj.p.val']), None)
    logfc_col = next((c for c in df.columns if c.lower() in ['logfc', 'log2foldchange']), None)
    pvalue_col = next((c for c in df.columns if c.lower() in ['pvalue', 'p.value']), None)
    geneid_col = next((c for c in df.columns if c.lower() in ['geneid', 'id', 'generalid']), None)
    symbol_col = next((c for c in df.columns if 'symbol' in c.lower()), None)

    if not all([padj_col, logfc_col]):
        print(f"Hata: Gerekli sütunlar (padj/logFC) bulunamadı.")
        return

    # 4. Filtreleme
    # padj < 0.05 olanlar
    df_sig = df[df[padj_col] < 0.05].copy()
    # logFC > 1 (Artan) ve logFC < -1 (Azalan)
    df_artan = df_sig[df_sig[logfc_col] > 1].copy()
    df_azalan = df_sig[df_sig[logfc_col] < -1].copy()
    df_tum = pd.concat([df_artan, df_azalan])

    # 5. Düzenleme ve Kaydetme Fonksiyonu
    def kaydet_excel(data, output_name):
        if data.empty:
            return

        # Sütun seçimi ve isimlendirme
        kolon_haritasi = {}
        if geneid_col: kolon_haritasi[geneid_col] = 'geneID'
        if padj_col: kolon_haritasi[padj_col] = 'padj'
        if pvalue_col: kolon_haritasi[pvalue_col] = 'pvalue'
        if logfc_col: kolon_haritasi[logfc_col] = 'logFC'
        if symbol_col: kolon_haritasi[symbol_col] = 'symbol'

        final_df = data[list(kolon_haritasi.keys())].rename(columns=kolon_haritasi)

        # logFC değerine göre küçükten büyüğe sırala
        final_df = final_df.sort_values(by='logFC', ascending=True)

        # Excel yazıcısını başlat
        writer = pd.ExcelWriter(output_name, engine='openpyxl')
        final_df.to_excel(writer, index=False, sheet_name='AnalizSonuc')

        workbook = writer.book
        worksheet = writer.sheets['AnalizSonuc']

        # Sadece padj sütunu için ondalıklı format (10 basamak)
        num_format = '0.0000000000'

        for idx, col_name in enumerate(final_df.columns):
            if col_name == 'padj':  # Sadece padj kontrolü
                for row in range(2, len(final_df) + 2):
                    worksheet.cell(row=row, column=idx + 1).number_format = num_format

        writer.close()

    # 6. İşlemleri Başlat
    kaydet_excel(df_artan, f"artan_genler{ek_isim}.xlsx")
    kaydet_excel(df_azalan, f"azalan_genler{ek_isim}.xlsx")
    kaydet_excel(df_tum, f"tum_genler{ek_isim}.xlsx")

    print(f"Başarıyla tamamlandı! Sadece 'padj' sütunu formatlandı.")


if __name__ == "__main__":
    analiz_yap()