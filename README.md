# Lotin-Kirill konvertor (Excel uchun)

**O'zbekcha** | [English](#english)

Excel'dagi o'zbekcha matnni lotin yozuvidan kirillga va kirilldan lotinga o'giradigan kichik qo'shimcha.

Muallif: **Begzodbek Akramjonov**

## Nima qila oladi

- Belgilangan kataklarni bir bosishda o'giradi: **Lotin → Kirill** yoki **Kirill → Lotin**. Shu buyruqlar sichqonchaning o'ng tugmasi menyusida ham bor.
- Faqat matnga tegadi. Formulalar, sonlar va sanalar o'zgarmaydi.
- Adashib bosilsa, **Ortga qaytarish** tugmasi yoki Ctrl+Z hammasini joyiga qaytaradi.
- Formula bilan ham ishlaydi: `=KIRILL(A1)` va `=LOTIN(A1)`.
- Ruschadan kirgan 1100 dan ortiq qiyin so'zni taniydi: dvigatel → двигатель, retsidivist → рецидивист, rayonlar → районлар.
- Avtomobil raqamlari (40 393 JBA), sayt manzillari, e-mail, rim raqamlari (II chorak) va A4, USD kabi belgilarni o'zgartirmaydi.
- Internetsiz ishlaydi, hajmi 65 KB atrofida.

## O'rnatish

1. [Releases](../../releases) bo'limidan `Konvertor_Excel.zip` ni yuklab oling va bitta papkaga chiqaring.
2. Excel'ni yoping va `Ornatish.bat` ni ishga tushiring.
3. Excel'ni oching, yuqorida **Lotin-Kirill** yorlig'i chiqadi.

Qo'lda o'rnatish ham mumkin: `Konvertor.xlam` faylini `%APPDATA%\Microsoft\Excel\XLSTART` papkasiga ko'chiring.

O'chirish uchun shu papkadan `Konvertor.xlam` ni o'chirib tashlang.

## Talablar

Windows va Microsoft Excel 2010 yoki undan yangisi. Excel 2007 da lenta ishlaydi, lekin o'ng tugma menyusidagi buyruqlar chiqmaydi.

## Manba kodi

Makrosning to'liq kodi `src/Konvertor.bas` faylida. Excel'ning o'zida Alt+F11 bosib ham ko'rish mumkin.

## Xatolar va takliflar

Noto'g'ri o'girilgan so'z ko'rsangiz, [Issues](../../issues) bo'limida yozib qoldiring: lotincha so'z va uning to'g'ri kirillchasi.

## Mualliflik va litsenziya

Dastur muallifi — Begzodbek Akramjonov. Dastur MIT litsenziyasi bilan, bitta qo'shimcha shart bilan tarqatiladi: dastur yoki uning o'zgartirilgan nusxasi tarqatilganda, muallif ismi asl dasturdagi joylarda (lentadagi "Muallif" bo'limi va "Yordam" oynasi) ko'rinib turishi kerak. To'liq matn `LICENSE` faylida, boshqa ochiq manbalar litsenziyasi `Litsenziyalar.txt` faylida.

---

<a name="english"></a>

# Latin-Cyrillic converter for Excel

[O'zbekcha](#lotin-kirill-konvertor-excel-uchun) | **English**

A small Excel add-in that converts Uzbek text between the Latin and Cyrillic alphabets.

Author: **Begzodbek Akramjonov**

## Features

- Converts the selected cells in one click: **Lotin → Kirill** (Latin to Cyrillic) or **Kirill → Lotin** (Cyrillic to Latin). Both commands are also in the right-click menu.
- Changes text only. Formulas, numbers and dates are left untouched.
- **Ortga qaytarish** (Undo) or Ctrl+Z reverts the last conversion.
- Worksheet functions: `=KIRILL(A1)` (to Cyrillic) and `=LOTIN(A1)` (to Latin).
- Knows over 1,100 tricky Russian loanwords: dvigatel → двигатель, retsidivist → рецидивист, rayonlar → районлар.
- Leaves car plate numbers (40 393 JBA), web addresses, e-mails, Roman numerals (II chorak) and codes such as A4 or USD unchanged.
- Works offline; about 65 KB in size.

## Installation

1. Download `Konvertor_Excel.zip` from [Releases](../../releases) and extract it into one folder.
2. Close Excel and run `Ornatish.bat`.
3. Open Excel — a **Lotin-Kirill** tab appears on the ribbon.

Manual installation: copy `Konvertor.xlam` into `%APPDATA%\Microsoft\Excel\XLSTART`.

To uninstall, delete `Konvertor.xlam` from that folder.

## Requirements

Windows and Microsoft Excel 2010 or later. In Excel 2007 the ribbon tab works, but the right-click menu commands do not appear.

## Source code

The full macro source is in `src/Konvertor.bas`. You can also view it inside Excel with Alt+F11.

## Bugs and suggestions

If a word is converted incorrectly, please open an [issue](../../issues) with the Latin word and its correct Cyrillic spelling.

## Authorship and license

Author: Begzodbek Akramjonov. The add-in is released under the MIT License with one additional condition: any distributed copy or modified version must keep the author's name visible where the original shows it (the "Muallif" group on the ribbon and the help window). See `LICENSE` for the full text and `Litsenziyalar.txt` for third-party notices.
