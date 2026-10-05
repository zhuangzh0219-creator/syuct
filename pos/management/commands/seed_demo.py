from django.core.management.base import BaseCommand

from pos.models import Customer, Product


class Command(BaseCommand):
    help = "POS画面の確認用サンプル商品・顧客を登録します"

    def handle(self, *args, **options):
        samples = [
            ("コーヒー", "DR-001", 450, 30),
            ("抹茶ラテ", "DR-002", 600, 24),
            ("チーズケーキ", "FD-001", 520, 16),
            ("季節のタルト", "FD-002", 680, 12),
            ("オリジナルマグカップ", "GO-001", 1200, 8),
        ]
        for name, sku, price, stock in samples:
            Product.objects.get_or_create(
                sku=sku,
                defaults={"name": name, "price": price, "stock": stock},
            )
        Customer.objects.get_or_create(name="山田 花子", defaults={"email": "hanako@example.jp", "phone": "090-1234-5678"})
        self.stdout.write(self.style.SUCCESS("サンプル商品と顧客を登録しました。"))
