"""Yandex Cloud Search API v2 — WordstatService.GetTop.
   Получает статистику поисковых запросов через Yandex Cloud gRPC.
"""
import grpc
from grpc_reflection.v1alpha.proto_reflection_descriptor_database import ProtoReflectionDescriptorDatabase
from google.protobuf.descriptor_pool import DescriptorPool
from google.protobuf import json_format
from google.protobuf.message_factory import GetMessageClass
import json

API_KEY = "AQVNyBi4wPgqGtaKV9-dvavzCH5TSaUNwBPvAErO"
FOLDER_ID = "b1gm97c74ca51sf4a5s2"

# Подключаемся к Search API gRPC endpoint
channel = grpc.secure_channel(
    'searchapi.api.cloud.yandex.net:443',
    grpc.ssl_channel_credentials(),
    options=[('grpc.primary_user_agent', 'yc-python-sdk')],
)

# Получаем дескрипторы через reflection
reflection_db = ProtoReflectionDescriptorDatabase(channel)
pool = DescriptorPool(reflection_db)

# Получаем классы сообщений
GetTopRequestClass = GetMessageClass(pool.FindMessageTypeByName('yandex.cloud.searchapi.v2.GetTopRequest'))
GetTopResponseClass = GetMessageClass(pool.FindMessageTypeByName('yandex.cloud.searchapi.v2.GetTopResponse'))
WordstatServiceStub = type('WordstatServiceStub', (), {})

# Metadata для авторизации
METADATA = [
    ('authorization', f'Api-Key {API_KEY}'),
    ('x-folder-id', FOLDER_ID),
]

def get_wordstat_top(phrase: str, num_phrases: int = 10) -> dict:
    """Получить статистику Wordstat для фразы."""
    request = GetTopRequestClass()
    request.phrase = phrase
    request.num_phrases = num_phrases
    request.folder_id = FOLDER_ID

    # Создаём универсальный unary-unary метод
    method_name = '/yandex.cloud.searchapi.v2.WordstatService/GetTop'

    # unary_unary(method, request_serializer, response_deserializer)(request, metadata=...)
    response = channel.unary_unary(
        method_name,
        request_serializer=lambda req: req.SerializeToString(),
        response_deserializer=lambda bs: GetTopResponseClass().parse(bs),
    )(request, metadata=METADATA)

    return json.loads(json_format.MessageToJson(response, preserving_proto_field_name=True))

# Список 7 запросов для Расцвет
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
for phrase, key in queries:
    print(f"\n=== {phrase} ===")
    try:
        r = get_wordstat_top(phrase, num_phrases=10)
        total_count = r.get('totalCount', 0) if isinstance(r, dict) else 0
        # Если total_count через строку (int64 как строка)
        if isinstance(total_count, str):
            total_count = int(total_count)
        print(f"  total_count: {total_count}")
        print(f"  results (первые 5): {r.get('results', [])[:5]}")
        results[key] = {'phrase': phrase, 'count': total_count, 'raw': r}
    except grpc.RpcError as e:
        print(f"  ERROR: {e.code()}: {e.details()}")
        results[key] = {'phrase': phrase, 'count': 0, 'error': f"{e.code()}: {e.details()}"}
    except Exception as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        results[key] = {'phrase': phrase, 'count': 0, 'error': str(e)}

# Итог
print("\n\n=== ИТОГ ===")
print(f"{'Тип запроса':<20} | {'Показов/мес':<12}")
print("-" * 40)
total = 0
for key, r in results.items():
    cnt = r.get('count', 0)
    total += cnt
    print(f"{key:<20} | {cnt:<12}")
print(f"{'ИТОГО':<20} | {total:<12}")

# Сохраняем
with open('/home/z/my-project/scripts/rascvet_wordstat.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2, default=str)
print("\nДанные сохранены в rascvet_wordstat.json")

channel.close()
