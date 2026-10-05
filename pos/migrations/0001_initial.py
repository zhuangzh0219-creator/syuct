from django.db import migrations, models
import django.db.models.deletion
import pos.models


class Migration(migrations.Migration):
    initial = True

    dependencies = []

    operations = [
        migrations.CreateModel(
            name="Customer",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("name", models.CharField(max_length=120, verbose_name="顧客名")),
                ("email", models.EmailField(blank=True, max_length=254, verbose_name="メールアドレス")),
                ("phone", models.CharField(blank=True, max_length=30, verbose_name="電話番号")),
                ("created_at", models.DateTimeField(auto_now_add=True, verbose_name="登録日時")),
            ],
            options={"verbose_name": "顧客", "verbose_name_plural": "顧客", "ordering": ["name"]},
        ),
        migrations.CreateModel(
            name="Product",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("name", models.CharField(max_length=120, verbose_name="商品名")),
                ("sku", models.CharField(max_length=40, unique=True, verbose_name="商品コード")),
                ("price", models.PositiveIntegerField(verbose_name="価格（税込円）")),
                ("stock", models.PositiveIntegerField(default=0, verbose_name="在庫数")),
                ("is_active", models.BooleanField(default=True, verbose_name="販売中")),
                ("created_at", models.DateTimeField(auto_now_add=True, verbose_name="登録日時")),
            ],
            options={"verbose_name": "商品", "verbose_name_plural": "商品", "ordering": ["name"]},
        ),
        migrations.CreateModel(
            name="Sale",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("receipt_number", models.CharField(default=pos.models.create_receipt_placeholder, max_length=24, unique=True, verbose_name="領収書番号")),
                ("status", models.CharField(choices=[("paid", "支払済み")], default="paid", max_length=12, verbose_name="状態")),
                ("subtotal", models.PositiveIntegerField(verbose_name="小計（税込円）")),
                ("total", models.PositiveIntegerField(verbose_name="合計（税込円）")),
                ("created_at", models.DateTimeField(auto_now_add=True, verbose_name="販売日時")),
                ("customer", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="sales", to="pos.customer", verbose_name="顧客")),
            ],
            options={"verbose_name": "販売", "verbose_name_plural": "販売", "ordering": ["-created_at"]},
        ),
        migrations.CreateModel(
            name="SaleItem",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("product_name", models.CharField(max_length=120, verbose_name="商品名")),
                ("unit_price", models.PositiveIntegerField(verbose_name="単価（税込円）")),
                ("quantity", models.PositiveIntegerField(verbose_name="数量")),
                ("line_total", models.PositiveIntegerField(verbose_name="金額（税込円）")),
                ("product", models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, related_name="sale_items", to="pos.product", verbose_name="商品")),
                ("sale", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="items", to="pos.sale", verbose_name="販売")),
            ],
            options={"verbose_name": "販売明細", "verbose_name_plural": "販売明細"},
        ),
        migrations.CreateModel(
            name="Payment",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("method", models.CharField(choices=[("cash", "現金"), ("paypay", "PayPay"), ("alipay", "支付宝 / Alipay"), ("card", "クレジットカード（デモ）")], max_length=16, verbose_name="支払方法")),
                ("amount", models.PositiveIntegerField(verbose_name="支払額（税込円）")),
                ("transaction_reference", models.CharField(max_length=40, unique=True, verbose_name="取引参照番号")),
                ("paid_at", models.DateTimeField(auto_now_add=True, verbose_name="支払日時")),
                ("sale", models.OneToOneField(on_delete=django.db.models.deletion.CASCADE, related_name="payment", to="pos.sale", verbose_name="販売")),
            ],
            options={"verbose_name": "決済", "verbose_name_plural": "決済", "ordering": ["-paid_at"]},
        ),
    ]
