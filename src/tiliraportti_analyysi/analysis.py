"""Read settlement rows and attribute fees using payment identifiers."""

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from io import BytesIO
from zipfile import ZipFile

from defusedxml.ElementTree import iterparse
from defusedxml.common import DefusedXmlException
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

REQUIRED = {"Maksun tunniste", "Joukkueet", "Maksuosuuden tyyppi", "Maksuosuus"}
UNALLOCATED = "Kohdistamatta"
MAX_UPLOAD_BYTES = 5 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 20 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 128
MAX_ROWS = 20_000  # Excludes the header.
MAX_COLUMNS = 64
MAX_CELLS = 200_000
MAX_XML_ELEMENTS = 1_000_000


def validate_workbook(data: bytes) -> None:
    """Bound ZIP expansion and XML work before handing untrusted data to openpyxl."""
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError("Tiedosto on liian suuri. Enimmäiskoko on 5 MiB.")
    try:
        with ZipFile(BytesIO(data)) as archive:
            members = archive.infolist()
            if (len(members) > MAX_ARCHIVE_MEMBERS
                    or len({member.filename for member in members}) != len(members)
                    or sum(member.file_size for member in members) > MAX_UNCOMPRESSED_BYTES):
                raise ValueError("Excelin purettu sisältö on liian suuri tai rakenne ei ole tuettu.")
            cells = elements = 0
            for member in members:
                if member.flag_bits & 1:
                    raise ValueError("Salattua Excel-tiedostoa ei tueta.")
                if not member.filename.endswith((".xml", ".rels")):
                    continue
                with archive.open(member) as source:
                    for _, element in iterparse(source, events=("end",), forbid_dtd=True):
                        elements += 1
                        if elements > MAX_XML_ELEMENTS:
                            raise ValueError("Excel sisältää liikaa XML-elementtejä.")
                        tag = element.tag.rsplit("}", 1)[-1]
                        if member.filename.startswith("xl/worksheets/"):
                            if tag == "c":
                                cells += 1
                                if cells > MAX_CELLS:
                                    raise ValueError("Excel sisältää liikaa soluja (enintään 200 000).")
                            if tag == "row" and int(element.get("r", "1")) > MAX_ROWS + 1:
                                raise ValueError("Excel sisältää liikaa rivejä (enintään 20 000).")
                            reference = element.get("ref") if tag == "dimension" else (
                                element.get("r") if tag == "c" else None)
                            if reference:
                                _, _, column, row = range_boundaries(reference)
                                if (column is None or row is None or column > MAX_COLUMNS
                                        or row > MAX_ROWS + 1):
                                    raise ValueError("Excelin enimmäiskoko on 20 000 riviä ja 64 saraketta.")
                        element.clear()
    except DefusedXmlException as exc:
        raise ValueError("Excelin XML-sisältö ei ole turvallinen. DTD:t ja entiteetit on estetty.") from exc
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Excelin rakenne on virheellinen tai sen XML-sisältö ei ole turvallinen.") from exc


@dataclass(frozen=True)
class Entry:
    row: int
    key: tuple[str, str, str]
    team: str
    kind: str
    amount: Decimal


@dataclass
class Result:
    summary: list[dict]
    details: list[dict]
    warnings: list[str]


def text(value) -> str:
    return "" if value is None else str(value).strip()


def money(value, row: int) -> Decimal:
    try:
        result = Decimal(text(value).replace("\u00a0", "").replace(" ", "").replace(",", "."))
        if (not result.is_finite() or result.copy_abs() > Decimal("1e12")
                or not -6 <= result.as_tuple().exponent <= 12):
            raise InvalidOperation
        return result
    except InvalidOperation:
        raise ValueError(f"Rivi {row}: Maksuosuus ei ole kelvollinen luku ({value!r}).") from None


def analyze(data: bytes) -> Result:
    validate_workbook(data)
    try:
        workbook = load_workbook(BytesIO(data), read_only=True, data_only=True, keep_links=False)
    except Exception as exc:
        raise ValueError("Tiedostoa ei voi lukea Excel-työkirjana. Valitse ehjä .xlsx-tiedosto.") from exc
    try:
        if "Tilitysrivit" not in workbook.sheetnames:
            raise ValueError("Excelistä puuttuu Tilitysrivit-välilehti.")
        sheet = workbook["Tilitysrivit"]
        # Do not trust worksheet dimensions: they can hide data or inflate iteration.
        sheet.reset_dimensions()
        rows = sheet.iter_rows(values_only=True, max_col=MAX_COLUMNS)
        headers = [text(value) for value in next(rows, ())]
        while headers and not headers[-1]:
            headers.pop()
        missing = REQUIRED - set(headers)
        if missing:
            raise ValueError("Puuttuvat sarakkeet: " + ", ".join(sorted(missing)))
        if len(headers) != len(set(headers)):
            raise ValueError("Otsikkorivillä on samannimisiä sarakkeita.")
        entries = []
        warnings = []
        for number, values in enumerate(rows, 2):
            if number > MAX_ROWS + 1:
                raise ValueError("Excel sisältää liikaa rivejä (enintään 20 000).")
            if all(value is None or text(value) == "" for value in values):
                continue
            row = dict(zip(headers, values))
            kind = text(row.get("Maksuosuuden tyyppi"))
            if kind not in {"Maksu", "Palvelumaksu", "Komissio"}:
                raise ValueError(f"Rivi {number}: tuntematon maksuosuuden tyyppi {kind!r}.")
            key = tuple(text(row.get(col)) for col in (
                "Organisaatio", "Tilityksen viitenumero", "Maksun tunniste"
            ))
            entries.append(Entry(number, key, text(row.get("Joukkueet")), kind,
                                 money(row.get("Maksuosuus"), number)))
        if not entries:
            raise ValueError("Tilitysrivit-välilehdellä ei ole analysoitavia rivejä.")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Excelin rivejä ei voi lukea. Valitse ehjä .xlsx-tiedosto.") from exc
    finally:
        workbook.close()

    teams = defaultdict(set)
    for entry in entries:
        if entry.kind == "Maksu" and entry.team and entry.key[2]:
            teams[entry.key].add(entry.team)
    totals = defaultdict(lambda: {"Maksu": Decimal(0), "Palvelumaksu": Decimal(0), "Komissio": Decimal(0)})
    payments = defaultdict(set)
    details = []
    for entry in entries:
        team = entry.team
        if not team:
            candidates = teams[entry.key] if entry.key[2] else set()
            if len(candidates) == 1:
                team = next(iter(candidates))
            else:
                team = UNALLOCATED
                warnings.append(f"Rivi {entry.row}: {entry.kind.lower()} jäi kohdistamatta "
                                f"(maksun tunniste {entry.key[2] or 'puuttuu'}).")
        totals[team][entry.kind] += entry.amount
        if entry.key[2]:
            payments[team].add(entry.key)
        details.append({"Excel-rivi": entry.row, "Joukkue": team,
                        "Maksun tunniste": entry.key[2], "Tyyppi": entry.kind,
                        "Maksuosuus (€)": entry.amount})
    summary = []
    for team, values in sorted(totals.items()):
        gross = values["Maksu"]
        service, commission = -values["Palvelumaksu"], -values["Komissio"]
        costs = service + commission
        summary.append({"Joukkue": team, "Maksuja (kpl)": len(payments[team]),
                        "Maksut (€)": gross, "Palvelumaksut (€)": service,
                        "Komissiot (€)": commission, "Kulut yhteensä (€)": costs,
                        "Palvelumaksut (%)": service / gross * 100 if gross else None,
                        "Komissiot (%)": commission / gross * 100 if gross else None,
                        "Kulut (%)": costs / gross * 100 if gross else None,
                        "Nettotilitys (€)": gross - costs})
    return Result(summary, details, warnings)
