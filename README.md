# CS2 Inventory Valuer (VBA)

Narzędzie oparte na języku **VBA (Excel)**, które pozwala na automatyczne pobieranie cen przedmiotów z gry Counter-Strike 2 bezpośrednio do arkusza kalkulacyjnego. 
Idealne dla traderów i kolekcjonerów chcących monitorować wartość swojego portfela.

## ⚠️ Ograniczenia API i Rate Limits
Program korzysta z oficjalnego API Steam (`market/priceoverview`), które posiada rygorystyczne limity zapytań (Rate Limits).

* **Zalecane zastosowanie:** Małe i średnie ekwipunki (do ok. 30-50 unikalnych przedmiotów).
* **Mechanizm zabezpieczający:** W kodzie zastosowano funkcję `Czekaj`, aby zminimalizować ryzyko blokady IP, jednak przy dużych kolekcjach Steam może tymczasowo przestać zwracać ceny.
* **Rozwiązanie:** Jeśli otrzymasz komunikat o limicie, odczekaj ok. 15 minut przed ponownym uruchomieniem.

### 🚀 Funkcje
* **Live Pricing:** Pobieranie aktualnych cen rynkowych dla pojedynczych skinów.
* **Full Inventory Scan:** Automatyczne obliczanie całkowitej wartości ekwipunku Steam.
* **Currency Support:** Przeliczanie wartości na wybraną walutę (zgodnie z API).

### 🛠️ Technologia
* **Język:** VBA (Visual Basic for Applications)
* **Źródło danych:** Steam Community Market API / Price APIs
* **Format:** Plik Excel (.xlsm)

### 📋 Jak używać?
1. Pobierz plik `.xlsm` z sekcji Releases.
2. Kliknij przycisk "Sprawdź Ekwipunek" i wpisz swój SteamID64.
3. Kliknij przycisk "Popraw Ceny", aby odświeżyć ceny ktorych API steama nie pobrało.
