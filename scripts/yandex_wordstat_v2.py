"""Yandex Cloud Search API v2 — WordstatService.GetTop
   Сбор данных по 7 запросам для бренда "Расцвет" (по аналогии с TITAN/Dali и т.д.)
"""
import grpc
from grpc_reflection.v1alpha.proto_reflection_descriptor_database import ProtoReflectionDescriptorDatabase
from google.protobuf.descriptor_pool import DescriptorPool
from google.protobuf import json_format
from google.protobuf.message_factory import GetMessageClass
import json
import time

API_KEY = "AQVNyBi4wPgqGtaKV9-dvavzCH5TSaUNwBPvAErO"
FOLDER_ID = "b1g0bo7majef8lq3iii2"

channel = grpc.secure_channel('searchapi.api.cloud.yandex.net:443', grpc.ssl_channel_credentials())
reflection_db = ProtoReflectionDescriptorDatabase(channel)
pool = DescriptorPool(reflection_db)

GetTopRequestClass = GetMessageClass(pool.FindMessageTypeByName('yandex.cloud.searchapi.v2.GetTopRequest'))
GetTopResponseClass = GetMessageClass(pool.FindMessageTypeByName('yandex.cloud.searchapi.v2.GetTopResponse'))

METADATA = [('authorization', f'Api-Key {API_KEY}'), ('x-folder-id', FOLDER_ID)]

def get_wordstat_top(phrase: str, num_phrases: int = 5) -> dict:
    req = GetTopRequestClass()
    req.phrase = phrase
    req.num_phrases = num_phrases
    req.folder_id = FOLDER_ID
    try:
        req.devices.append(1)  # DEVICE_ALL
    except Exception:
        pass

    raw_bytes = channel.unary_unary(
        '/yandex.cloud.searchapi.v2.WordstatService/GetTop',
        request_serializer=lambda r: r.SerializeToString(),
        response_deserializer=lambda bs: bs,
    )(req, metadata=METADATA)

    resp = GetTopResponseClass()
    resp.ParseFromString(raw_bytes)
    return json.loads(json_format.MessageToJson(resp, preserving_proto_field_name=True))


# === 7 запросов по бренду "Расцвет" (аналог методики для других брендов) ===
queries = [
    ("Расцвет краска",         "краска"),
    ("Расцвет грунт-эмаль",    "грунт-эмаль"),
    ("Расцвет 3 в 1",          "3 в 1"),
    ("Расцвет по металлу",     "по металлу"),
    ("Расцвет по ржавчине",    "по ржавчине"),
    ("Расцвет молотковая",     "молотковая"),
    ("Расцвет быстросохнущая", "быстросохнущая"),
]

results = {}
print(f"{'Тип запроса':<20} | {'Запрос':<28} | {'Показов/мес':<12}")
print("-" * 70)
total = 0
for phrase, key in queries:
    try:
        r = get_wordstat_top(phrase, 5)
        count = int(r.get('total_count', 0))
        # Также берём результаты (related phrases)
        related = r.get('results', [])
        print(f"{key:<20} | {phrase:<28} | {count:<12}")
        results[key] = {
            'phrase': phrase,
            'count': count,
            'related': related[:5],
        }
        total += count
    except grpc.RpcError as e:
        print(f"{key:<20} | {phrase:<28} | ERROR: {e.code()}: {e.details()[:100]}")
        results[key] = {'phrase': phrase, 'count': 0, 'error': str(e.details())}
    except Exception as e:
        print(f"{key:<20} | {phrase:<28} | ERROR: {e}")
        results[key] = {'phrase': phrase, 'count': 0, 'error': str(e)}
    time.sleep(0.5)  # мягкий rate limit

print("-" * 70)
print(f"{'ИТОГО':<20} | {'':<28} | {total:<12}")

# Сохраняем результат
with open('/home/z/my-project/scripts/rascvet_wordstat.json', 'w', encoding='utf-8') as f:
    json.dump({
        'brand': 'Расцвет',
        'queries': results,
        'total_7_requests': total,
        'timestamp': '2026-09-14',
    }, f, ensure_ascii=False, indent=2, default=str)

print(f"\n✅ Данные сохранены в scripts/rascvet_wordstat.json")
print(f"Сумма 7 запросов: {total}")
print(f"Для сравнения: TITAN = 1 430, Correx = 1 185")

channel.close()
