# VBA-Automatic-report-generation
Business case:
---------------------
Należy stworzyć proste narzędzie w Excelu z wykorzystaniem Power Query oraz VBA. Narzędzie powinno pobierać dane źródłowe, następnie obrabiać dane – połączyć w jedną tabelę zbiorczą zawierającą faktury sprzedażowe, nazwy klientów i przypisanych użytkowników. Następnie należy podzielić tabelę zbiorczą na oddzielne pliki .xlsx dla każdego z użytkowników.

Opis plików źródłowych:
Jest to wyciąg danych z systemu CRM. W folderze invoices znajdują się pliki z fakturami sprzedażowymi z poszczególnych lat. W pliku Excel o nazwie "clients_and_users" znajdują się 2 tabele:
- clients - tabela klientów
- users - tabela użytkowników

Wszystkie tabele są ze sobą powiązane za pomocą kluczy.

Całe narzędzie (zapytanie Power Query oraz makro VBA) powinno znajdować się w jednym pliku.

1. Stwórz mechanizm pobierania plików źródłowych za pomocą Power Query.

a) Zaprojektuj w Power Query pobieranie faktur sprzedażowych z folderu invoices, przy założeniu, że w kolejnych latach będą dochodzić kolejne faktury (kolejne pliki csv np. za rok 2020, 2021 itd.). Pliki źródłowe csv zawsze mają taką samą strukturę.

b) Zaprojektuj w Power Query pobieranie danych dotyczących klientów oraz użytkowników z pliku Excel o nazwie: "clients_and_users.xlsx". Mechanizm powinien przewidywać, że będą dochodzić kolejne rekordy z klientami oraz użytkownikami.

2. Stwórz tabelę zbiorczą. Połącz dane źródłowe tak, aby plik wynikowy miał strukturę analogiczną jak w pliku: PlikWynikowyWzor.xlsx. Chodzi o to, aby do tabeli faktur wstawić dodatkowe 4 kolumny: client_name, id_user, user_login, user_email. Na zielono zaznaczono 4 kolumny, które należy dodać do tabeli z fakturami.
 

3. Za pomocą makra VBA podziel tabelę zbiorczą na oddzielne pliki .xlsx. Pliki należy podzielić tak, aby każdy z użytkowników otrzymał oddzielny plik z przypisanymi do niego rekordami (fakturami). Przykładowo user1 powinien otrzymać plik z fakturami przypisanymi do user1.

Zadbaj, aby narzędzie było user friendly. Nie wymagamy tworzenia User Forms, wystarczy użycie przycisków i odpowiednie, proste zaprojektowanie arkusza, ale w taki sposób, aby narzędzie działało na komputerze użytkownika końcowego. Zakładamy, że osoba, która ma korzystać z narzędzia w Excelu nie zna Power Query oraz VBA. Jest to przeciętny użytkownik Excela.
