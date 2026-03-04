"""Обгортка над API Нової Пошти v2."""
import requests


NP_API_URL = "https://api.novaposhta.ua/v2.0/json/"


class NovaPoshtaError(Exception):
    pass


class NovaPoshtaAPI:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self._city_cache: dict[str, str] = {}
        self._warehouse_cache: dict[tuple, str] = {}

    def _call(self, model: str, method: str, props: dict) -> dict:
        payload = {
            "apiKey": self.api_key,
            "modelName": model,
            "calledMethod": method,
            "methodProperties": props,
        }
        resp = requests.post(NP_API_URL, json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if not data.get("success"):
            errors = ", ".join(data.get("errors", ["невідома помилка"]))
            raise NovaPoshtaError(f"{model}.{method}: {errors}")
        return data

    def get_city_ref(self, city_name: str) -> str:
        """Отримати Ref міста за назвою (з кешем)."""
        key = city_name.strip().lower()
        if key in self._city_cache:
            return self._city_cache[key]

        data = self._call("Address", "getCities", {"FindByString": city_name})
        items = data.get("data", [])
        if not items:
            raise NovaPoshtaError(f"Місто не знайдено: {city_name!r}")
        ref = items[0]["Ref"]
        self._city_cache[key] = ref
        return ref

    def get_warehouse_ref(self, city_ref: str, warehouse_number: str) -> str:
        """Отримати Ref відділення за номером у місті (з кешем)."""
        cache_key = (city_ref, str(warehouse_number))
        if cache_key in self._warehouse_cache:
            return self._warehouse_cache[cache_key]

        data = self._call(
            "Address",
            "getWarehouses",
            {"CityRef": city_ref, "WarehouseId": str(warehouse_number)},
        )
        items = data.get("data", [])
        if not items:
            raise NovaPoshtaError(
                f"Відділення #{warehouse_number} не знайдено у місті {city_ref}"
            )
        ref = items[0]["Ref"]
        self._warehouse_cache[cache_key] = ref
        return ref

    def create_counterparty(self, first_name: str, last_name: str, middle_name: str, phone: str) -> dict:
        """Створити/знайти контакт отримувача.
        Повертає dict з ключами: Ref, ContactPersons[0].Ref
        """
        data = self._call(
            "Counterparty",
            "save",
            {
                "FirstName": first_name,
                "MiddleName": middle_name,
                "LastName": last_name,
                "Phone": phone,
                "Email": "",
                "CounterpartyType": "PrivatePerson",
                "CounterpartyProperty": "Recipient",
            },
        )
        item = data["data"][0]
        counterparty_ref = item["Ref"]
        contact_ref = item["ContactPerson"]["data"][0]["Ref"]
        return {"counterparty_ref": counterparty_ref, "contact_ref": contact_ref}

    def find_ttn_by_internal_number(self, internal_number: str) -> str | None:
        """Шукає існуючу ТТН за InternalNumber (ID угод).
        Повертає IntDocNumber якщо знайдено, інакше None.
        """
        try:
            data = self._call(
                "InternetDocument",
                "getDocumentList",
                {"InternalNumber": internal_number, "GetFullList": "1"},
            )
            items = data.get("data", [])
            if items:
                return items[0].get("IntDocNumber")
        except NovaPoshtaError:
            pass
        return None

    def create_ttn(self, params: dict) -> str:
        """Створити ТТН. Повертає номер накладної (IntDocNumber)."""
        data = self._call("InternetDocument", "save", params)
        return data["data"][0]["IntDocNumber"]

    # === Методи для setup_refs.py ===

    def get_counterparties(self, prop: str = "Sender") -> list[dict]:
        data = self._call(
            "Counterparty",
            "getCounterparties",
            {"CounterpartyProperty": prop, "Page": "1"},
        )
        return data.get("data", [])

    def get_counterparty_contact_persons(self, counterparty_ref: str) -> list[dict]:
        data = self._call(
            "Counterparty",
            "getCounterpartyContactPersons",
            {"Ref": counterparty_ref},
        )
        return data.get("data", [])

    def get_counterparty_addresses(self, counterparty_ref: str) -> list[dict]:
        data = self._call(
            "Counterparty",
            "getCounterpartyAddresses",
            {"Ref": counterparty_ref, "CounterpartyProperty": "Sender"},
        )
        return data.get("data", [])
