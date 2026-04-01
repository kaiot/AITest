"""
Generate self-signed SSL certificates for local HTTPS.

Chrome requires HTTPS for microphone access, and WSS for WebSocket.
Run this once during setup. Generates cert.pem and key.pem.
"""

import sys
from pathlib import Path


def generate():
    try:
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509 import DNSName, IPAddress
        import ipaddress
        import datetime
    except ImportError:
        print("Installing cryptography...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "cryptography"])
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509 import DNSName, IPAddress
        import ipaddress
        import datetime

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "JARVIS Local"),
    ])

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.utcnow())
        .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=3650))
        .add_extension(
            x509.SubjectAlternativeName([
                DNSName("localhost"),
                IPAddress(ipaddress.IPv4Address("127.0.0.1")),
            ]),
            critical=False,
        )
        .sign(key, hashes.SHA256())
    )

    out_dir = Path(__file__).parent
    key_path = out_dir / "key.pem"
    cert_path = out_dir / "cert.pem"

    key_path.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption(),
        )
    )
    cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))

    print(f"Generated: {cert_path}")
    print(f"Generated: {key_path}")
    print()
    print("NOTE: Chrome will warn about the self-signed cert.")
    print("To trust it: open Chrome -> visit https://localhost:8340 -> Advanced -> Proceed.")
    print("Or run setup.bat which will open that page for you.")


if __name__ == "__main__":
    generate()
