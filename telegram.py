from telethon.sync import TelegramClient, events  # Telethon çağırıldı
from openpyxl import load_workbook  # openpyxl (excel kontrol kütüphanesi)
import os  # dosya silme, kontrol işlemleri
from datetime import datetime, date, timedelta  # Zaman kontrolü için kütüphane
from locale import setlocale, LC_ALL  # Zaman kütüphanesi türkçe çalışsın diye.
from const import api_id, api_hash

# Türkçe yapma
try:
    setlocale(LC_ALL, 'tr_TR')
except:
    setlocale(LC_ALL, 'tr_TR.utf8')

with TelegramClient('name', api_id, api_hash) as client:
    @client.on(events.NewMessage(pattern="plan"))  # Eğer "plan" yazılmışsa
    async def plan(event):
        await excelden_veri_getir(event)


    @client.on(events.NewMessage())
    async def plandoldur(event): # eğer gönderilen herhangi bir mesajda döküman varsa kaydet
        if event.media:
            await client.download_media(event.media, "telegramexceli.xlsx")


    @client.on(events.NewMessage(pattern="aylik")) # Eğer "aylık" yazılmışsa
    async def plan(event):
        await excelden_veri_getir(event, True)


    async def excelden_veri_getir(event, tum_gunler_mi=False):
        if not os.path.exists("telegramexceli.xlsx"):
            return
            # bu kod excel dosyası olduğunda devam edecek
        excel_dosyasi = load_workbook("telegramexceli.xlsx")
        sayfa1 = excel_dosyasi.active

        # entity = client.get_entity("andromeda606")
        yarin = (datetime.now() + timedelta(days=1)).day

        for a, b in zip(sayfa1["A"], sayfa1["B"]):
            if type(a.value) == str:
                continue
            if (a.value == yarin or a.value == yarin + 1 or a.value == yarin + 2) or tum_gunler_mi:
                bugun = datetime.now()
                ayin_basi = date(bugun.year, bugun.month, 1)
                print(a.value)
                tarih = ayin_basi + timedelta(days=a.value - 1)
                if b.value is None:
                    await event.reply(
                        "Ayın " + str(a.value) + ".günü (" + datetime.strftime(tarih, '%A') + ") boş")
                    continue
                veri = ""
                for yapilacak in b.value.split("***"):
                    veri += yapilacak.strip() + "\n"
                await event.reply(
                    "Ayın " + str(a.value) + ".gününde (" + datetime.strftime(tarih, '%A') + ");\n" + veri)

    client.run_until_disconnected()
