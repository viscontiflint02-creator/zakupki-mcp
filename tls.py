"""Доверять сертификатам из хранилища Windows.

Корпоративная сеть проверяет HTTPS и подменяет сертификаты сайтов своим корневым.
Windows ему доверяет, а Python (certifi, httplib2) — нет: «UnknownIssuer» при обращении к
Google и сайтам с ценами. truststore заставляет ssl проверять сертификаты через ОС.
"""


def use_system_certificates() -> None:
    try:
        import truststore
    except ImportError:  # старая установка без truststore — работаем как раньше
        return
    truststore.inject_into_ssl()
