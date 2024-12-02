import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# Ürün bilgilerini saklamak için bir liste
products = []
# eski ürün verileri için liste
old_products = []

# Excel dosyasını yükleme
def load_excel():
    global old_products
    file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Dosyası", "*.xlsx")])
    if not file_path:
        return

    try:
        # Excel dosyasını oku
        old_products = pd.read_excel(file_path)

        # url sütunundaki geçersiz url' leri kontrol et
        if 'url' not in old_products.columns:
            messagebox.showerror("Hata", "Excel dosyasında 'url' sütunu bulunamadı.")
            return
        
        # Geçersiz url'leri ayıkla ve uyarı ver
        valid_urls = old_products[old_products['url'].apply(lambda x: isinstance(x, str) and x.startswith('http'))]
        invalid_urls = old_products[~old_products.index.isin(valid_urls.index)]

        if not invalid_urls.empty:
            invalid_list = invalid_urls['url'].tolist()
            messagebox.showwarning(
                "Uyarı", 
                f"Geçersiz URL'ler bulundu ve atlandı:\n{', '.join([str(url) for url in invalid_list])}"
            )

        # Geçerli url'leri kullanmaya devam et
        old_products = valid_urls.reset_index(drop=True)
        display_excel_data_in_gui()

    except Exception as e:
        messagebox.showerror("Hata", f"Excel dosyası yüklenirken bir hata oluştu: {e}")


# Excel verilerini Tkinter' da göster
def display_excel_data_in_gui():
    if not old_products.empty:
        product_list.delete(1.0, tk.END)  # Önceki içerikleri temizle

        for _, row in old_products.iterrows():
            product_list.insert(tk.END, f"Ürün Adı: {row['name']}\nFiyat: {row['price']}\nURL: {row['url']}\n\n")
    else:
        messagebox.showerror("Hata", "Excel dosyasında veri yok!")

# Ürün bilgilerini çekme ve listeleme
def scrape_and_show():
    global products  # Ürünleri globalde tutalım
    max_products = int(entry_max_products.get())

    # Selenium WebDriver 
    driver = webdriver.Chrome()
    driver.maximize_window()

    # url'ye git
    url = "https://www.zara.com/tr/tr/man-all-products-l7465.html?v1=2443335"
    driver.get(url)

    products = []

    try:
        while len(products) < max_products:
            try:
                
                product_containers = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'product-grid-product__data')]"))
                )

                for container in product_containers:
                    try:
                        # Ürün adı bilgisi
                        product_name_element = container.find_element(By.XPATH, ".//a[contains(@class, 'product-link _item product-grid-product-info__name link')]//h2")
                        product_name = product_name_element.text.strip()

                        # Ürün fiyatı bilgisi
                        product_price_element = container.find_element(By.XPATH, ".//span[@class='money-amount__main']")
                        product_price = product_price_element.text.strip()

                        # Ürün URL'si
                        product_url_element = container.find_element(By.XPATH, ".//a[contains(@class, 'product-link _item product-grid-product-info__name link')]")
                        product_url = product_url_element.get_attribute("href")

                        # Ürün bilgilerini listeye ekle
                        products.append({
                            'name': product_name,
                            'price': product_price,
                            'url': product_url
                        })

                        # İstenilen sayıya ulaşıldıysa döngüyü kır
                        if len(products) >= max_products:
                            break
                    except StaleElementReferenceException:
                        continue

            except TimeoutException:
                messagebox.showerror("Hata", "Ürünler yüklenemedi!")
                break

            # Sayfa kaydırma işlemi
            last_height = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            WebDriverWait(driver, 5).until(lambda d: d.execute_script("return document.body.scrollHeight") > last_height)

    finally:
        driver.quit()

    # Ürün bilgilerini arayüzde göster
    product_list.delete(1.0, tk.END)  # Önceki içerikleri temizle
    for product in products:
        product_list.insert(tk.END, f"Ürün Adı: {product['name']}\nFiyat: {product['price']}\nURL: {product['url']}\n\n")

    # Ürünleri Excel dosyasına kaydetmek için kullanıcıya seçenek sun
    if messagebox.askyesno("Excel Kaydet", "Ürün verilerini Excel dosyasına kaydetmek ister misiniz?"):
        save_to_excel()

# Excel'e kaydetme fonksiyonu
def save_to_excel():
    global products
    if products:
        df = pd.DataFrame(products)
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Dosyası", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Başarılı", "Veriler başarıyla kaydedildi.")
    else:
        messagebox.showerror("Hata", "Kaydedilecek veri yok.")

# Fiyat değişikliklerini kontrol etme fonksiyonu
def check_price_changes():
    if old_products.empty:
        messagebox.showerror("Hata", "Lütfen önce Excel dosyasını yükleyin.")
        return

    try:
        driver = webdriver.Chrome()
        driver.maximize_window()
    except Exception as e:
        messagebox.showerror("Hata", f"WebDriver başlatılamadı: {e}")
        return

    changes_detected = False
    message = "Fiyat değişimleri:\n"

    try:
        for _, row in old_products.iterrows():
            url = row['url']  # Excel dosyasındaki url

            if not isinstance(url, str) or not url.startswith("http"):
                print(f"Geçersiz URL atlandı: {url}")
                continue  # Geçersiz URL'yi atla

            try:
                driver.get(url)
                
                # Ürünün fiyatını yeni sayfada kontrol et
                product_price_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='main']/article/div/div[1]/div[2]/div/div[1]/div[2]/div/span/span/span/div/span"))
                )
                new_price = product_price_element.text.strip()

                # Eski fiyatı ve yeni fiyatı karşılaştır
                if new_price != str(row['price']).strip():
                    changes_detected = True
                    message += f"Ürün: {row['name']}\nEski Fiyat: {row['price']} -> Yeni Fiyat: {new_price}\n\n"

            except TimeoutException:
                message += f"Ürün: {row['name']}\nFiyat alınamadı (Timeout).\n\n"
            except Exception as e:
                message += f"Ürün: {row['name']}\nBir hata oluştu: {e}\n\n"

    finally:
        driver.quit()

    if changes_detected:
        messagebox.showinfo("Fiyat Değişiklikleri", message)
    else:
        messagebox.showinfo("Fiyat Değişiklikleri", "Fiyat değişikliği bulunamadı.")


# ARAYÜZ
def run_gui():
    global product_list

    root = tk.Tk()
    root.title("Zara Ürün Fiyat Takibi")
    root.geometry("600x600")

    # Ürün sayısını girecek alan
    tk.Label(root, text="Kaç ürün çekmek istersiniz?").pack(pady=5)
    global entry_max_products
    entry_max_products = tk.Entry(root)
    entry_max_products.pack(pady=10)
    entry_max_products.insert(0, "5")  # Varsayılan değer 5

    # Fiyat verilerini çekme butonu
    fetch_button = tk.Button(root, text="Fiyat Verilerini Çek", command=scrape_and_show)
    fetch_button.pack(pady=10)

    # Ürünleri gösterecek yer
    product_list = tk.Text(root, height=15, width=70)
    product_list.pack(pady=10)

    # Excel dosyasını yükle butonu
    load_button = tk.Button(root, text="Excel Dosyasını Yükle", command=load_excel)
    load_button.pack(pady=10)

    # Fiyat değişimini kontrol etme 
    check_changes_button = tk.Button(root, text="Fiyat Değişimlerini Gör", command=check_price_changes)
    check_changes_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
