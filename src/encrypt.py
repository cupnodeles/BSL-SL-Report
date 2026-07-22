# src/encrypt.py
# BPI BL SL Automation — File Open Password Encryption & Decryption
# Password: SPM*123
# Uses msoffcrypto-tool

import io
import msoffcrypto
import logging

logger = logging.getLogger("BPI_BL_SL")

PASSWORD = "SPM*123"


def encrypt_file(file_bytes: io.BytesIO) -> io.BytesIO:
    """
    Applies file open password 'SPM*123' to an Excel BytesIO object.
    Returns encrypted BytesIO ready for download.
    """
    logger.info("Encrypting output file with password...")
    try:
        file_bytes.seek(0)
        encrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file_bytes)
        office_file.encrypt(PASSWORD, encrypted)
        encrypted.seek(0)
        logger.info("File encrypted successfully.")
        return encrypted
    except Exception as e:
        logger.error(f"Encryption failed: {e}")
        raise


def decrypt_file(file_bytes: io.BytesIO) -> io.BytesIO:
    """
    Decrypts a password-protected Excel BytesIO object using 'SPM*123'.
    Returns decrypted BytesIO ready for processing.
    Used when uploading previously encrypted output files as templates.
    """
    logger.info("Decrypting input file with password...")
    try:
        file_bytes.seek(0)
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file_bytes)

        # Check if file is actually encrypted
        if not office_file.is_encrypted():
            logger.info("File is not encrypted. Returning as-is.")
            file_bytes.seek(0)
            return file_bytes

        office_file.load_key(password=PASSWORD)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        logger.info("File decrypted successfully.")
        return decrypted
    except Exception as e:
        logger.error(f"Decryption failed: {e}")
        raise


def is_encrypted(file_bytes: io.BytesIO) -> bool:
    """
    Checks if a file is encrypted/password protected.
    Returns True if encrypted, False otherwise.
    """
    try:
        file_bytes.seek(0)
        office_file = msoffcrypto.OfficeFile(file_bytes)
        result = office_file.is_encrypted()
        file_bytes.seek(0)
        return result
    except Exception:
        file_bytes.seek(0)
        return False