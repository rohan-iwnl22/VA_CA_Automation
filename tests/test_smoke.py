"""Test smoke test for package import."""

def test_import_package():
    import va_ca_automation

    assert va_ca_automation.__version__ == "0.1.0"
