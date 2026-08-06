import os
import requests
import time
from openpyxl import Workbook
from openpyxl.styles import Font

# ==========================================================
# CONFIGURACIÓN
# ==========================================================

API_KEY = ""
ORG_ID = ""
SSIDS = ["Escaners","IoT-Devices"]

NETWORK_FILE = "networks.txt"
EXCLUDED_MACS_FILE = "Macexcluidas.txt"

INTERVALO_SEGUNDOS = 60  # Pausa entre ejecuciones

# Palabras a buscar
KEYWORDS = [
    "iphone",
    "oppo",
    "honor",
    "zte",
    "redmi",
    "moto-",
    "Xiaomi",
    "Reno",
    "OPPO-",
    "Xiaomi Communications",
    "motorola",
    "Apple iPhone",
    "Apple",
    "iPad",
    "iPhone",
]

headers = {
    "X-Cisco-Meraki-API-Key": API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json"
}

# ==========================================================
# BUCLE CONTINUO
# ==========================================================

ciclo = 1

try:
    while True:
        timestamp_inicio = time.strftime('%Y-%m-%d %H:%M:%S')
        print("\n" + "=" * 90)
        print(f"🔄 INICIANDO CICLO #{ciclo} - [{timestamp_inicio}]")
        print("=" * 90)

        # ------------------------------------------------------
        # 1. LEER ARCHIVO DE MACs EXCLUIDAS EN CADA CICLO
        # ------------------------------------------------------
        EXCLUDED_MACS = []
        if os.path.exists(EXCLUDED_MACS_FILE):
            with open(EXCLUDED_MACS_FILE, "r", encoding="utf-8") as file:
                EXCLUDED_MACS = [line.strip().lower() for line in file if line.strip()]
            print(f"🛡️ MACs excluidas cargadas: {len(EXCLUDED_MACS)}")
        else:
            print(f"⚠️ El archivo '{EXCLUDED_MACS_FILE}' no existe. Se continuará sin lista de exclusión.")

        # ------------------------------------------------------
        # 2. LEER ARCHIVO DE NETWORKS EN CADA CICLO
        # ------------------------------------------------------
        if not os.path.exists(NETWORK_FILE):
            print(f"❌ Error: El archivo '{NETWORK_FILE}' no existe. Reintentando en {INTERVALO_SEGUNDOS}s...")
            time.sleep(INTERVALO_SEGUNDOS)
            ciclo += 1
            continue

        with open(NETWORK_FILE, "r", encoding="utf-8") as file:
            network_ids = [line.strip() for line in file if line.strip()]

        print(f"🌐 NETWORKS A PROCESAR: {len(network_ids)}")

        # ------------------------------------------------------
        # 3. PREPARAR EXCEL PARA ESTE CICLO
        # ------------------------------------------------------
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"

        ws.append([
            "NetworkID",
            "SSID",
            "MAC",
            "Nombre",
            "IP",
            "Fabricante",
            "Sistema Operativo",
            "Predicción",
            "Coincidencia",
            "Política Actual",
            "Resultado"
        ])

        for cell in ws[1]:
            cell.font = Font(bold=True)

        total_detectados = 0
        total_bloqueados = 0
        total_omitidos = 0
        total_excluidas = 0

        # ------------------------------------------------------
        # 4. PROCESAR CADA NETWORK
        # ------------------------------------------------------
        for NETWORK_ID in network_ids:

            clients_url = f"https://api.meraki.com/api/v1/networks/{NETWORK_ID}/clients"
            params = {
                "timespan": 43200,      # Últimas 12 horas
                "perPage": 1000
            }

            response = requests.get(
                clients_url,
                headers=headers,
                params=params
            )

            if response.status_code != 200:
                print(f"❌ Error obteniendo clientes para Network: {NETWORK_ID}")
                print(f"   Status: {response.status_code} | Respuesta: {response.text}")
                continue

            clients = response.json()

            for c in clients:

                # Filtro por SSID
                if c.get("ssid") not in SSIDS:
                    continue

                # Detectar dispositivos por palabras clave
                device_info = " ".join([
                    str(c.get("os", "")),
                    str(c.get("manufacturer", "")),
                    str(c.get("description", "")),
                    str(c.get("dhcpHostname", "")),
                    str(c.get("deviceTypePrediction", ""))
                ]).lower()

                if not any(keyword.lower() in device_info for keyword in KEYWORDS):
                    continue

                total_detectados += 1

                mac = c.get("mac")
                if not mac:
                    continue

                # Exclusión por lista blanca
                if mac.lower() in EXCLUDED_MACS:
                    print(f"🛡️ MAC Excluida ({EXCLUDED_MACS_FILE}): {mac} (Se omite bloqueo)")
                    total_excluidas += 1
                    continue

                # Consultar política actual
                policy_url = f"https://api.meraki.com/api/v1/networks/{NETWORK_ID}/clients/{mac}/policy"
                policy_response = requests.get(policy_url, headers=headers)

                if policy_response.status_code != 200:
                    print(f"❌ Error consultando política de la MAC: {mac}")
                    print(f"   Status: {policy_response.status_code} | Respuesta: {policy_response.text}")
                    continue

                current_policy = policy_response.json().get("devicePolicy", "")

                # Si ya está bloqueado, se omite silenciosamente
                if str(current_policy).lower() == "blocked":
                    total_omitidos += 1
                    continue

                # Cambiar estado a Blocked
                nombre = c.get("description") or c.get("dhcpHostname") or "Sin nombre"
                coincidencia = ", ".join([k for k in KEYWORDS if k.lower() in device_info])

                print("\n" + "-" * 80)
                print(f"Network ID       : {NETWORK_ID}")
                print(f"MAC              : {mac}")
                print(f"Nombre           : {nombre}")
                print(f"IP               : {c.get('ip')}")
                print(f"Fabricante       : {c.get('manufacturer')}")
                print(f"OS               : {c.get('os')}")
                print(f"Predicción       : {c.get('deviceTypePrediction')}")
                print(f"Coincidencia     : {coincidencia}")
                print(f"Política anterior: {current_policy}")

                payload = {"devicePolicy": "Blocked"}
                update = requests.put(policy_url, headers=headers, json=payload)

                if update.status_code in (200, 201):
                    print("✅ Estado cambiado exitosamente de NORMAL a BLOCKED")
                    total_bloqueados += 1
                    resultado = "Bloqueado"

                    # Agregar a Excel únicamente si el bloqueo fue exitoso
                    ws.append([
                        NETWORK_ID,
                        c.get("ssid"),
                        mac,
                        nombre,
                        c.get("ip"),
                        c.get("manufacturer"),
                        c.get("os"),
                        c.get("deviceTypePrediction"),
                        coincidencia,
                        current_policy,
                        resultado
                    ])
                else:
                    print("❌ Error bloqueando equipo")
                    print(f"   Status: {update.status_code} | Respuesta: {update.text}")

                # Evitar rate limits de la API
                time.sleep(0.3)

        # ------------------------------------------------------
        # 5. RESUMEN DEL CICLO Y GUARDADO CONDICIONAL DE EXCEL
        # ------------------------------------------------------
        print("\n" + "-" * 90)
        print(f"RESUMEN CICLO #{ciclo}")
        print("-" * 90)
        print(f"Dispositivos detectados : {total_detectados}")
        print(f"Nuevos bloqueados       : {total_bloqueados}")
        print(f"Ya bloqueados (omitidos): {total_omitidos}")
        print(f"MACs excluidas          : {total_excluidas}")

        if total_bloqueados > 0:
            nombre_excel = f"Resultados_Meraki_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb.save(nombre_excel)
            print(f"📄 Excel generado       : {nombre_excel}")
        else:
            print("ℹ️ No se realizaron cambios a 'Blocked' en este ciclo. No se generó archivo Excel.")

        print("=" * 90)
        print(f"⏳ Esperando {INTERVALO_SEGUNDOS} segundos para el siguiente ciclo...")
        
        ciclo += 1
        time.sleep(INTERVALO_SEGUNDOS)

except KeyboardInterrupt:
    print("\n\n🛑 Ejecución detenida manualmente por el usuario (Ctrl + C). ¡Hasta luego!")
