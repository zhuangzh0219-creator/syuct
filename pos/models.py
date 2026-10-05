import uuid

from django.db import models


def create_receipt_placeholder():
    return f"T-{uuid.uuid4().hex[:12]}"


class Product(models.Model):
    name = models.CharField("商品名", max_length=120)
    sku = models.CharField("商品コード", max_length=40, unique=True)
    price = models.PositiveIntegerField("価格（税込円）")
    stock = models.PositiveIntegerField("在庫数", default=0)
    is_active = models.BooleanField("販売中", default=True)
    created_at = models.DateTimeField("登録日時", auto_now_add=True)

    class Meta:
        ordering = ["name"]
        verbose_name = "商品"
        verbose_name_plural = "商品"

    def __str__(self):
        return self.name


class Customer(models.Model):
    name = models.CharField("顧客名", max_length=120)
    email = models.EmailField("メールアドレス", blank=True)
    phone = models.CharField("電話番号", max_length=30, blank=True)
    created_at = models.DateTimeField("登録日時", auto_now_add=True)

    class Meta:
        ordering = ["name"]
        verbose_name = "顧客"
        verbose_name_plural = "顧客"

    def __str__(self):
        return self.name


class Sale(models.Model):
    class Status(models.TextChoices):
        PAID = "paid", "支払済み"

    receipt_number = models.CharField("領収書番号", max_length=24, unique=True, default=create_receipt_placeholder)
    customer = models.ForeignKey(Customer, verbose_name="顧客", on_delete=models.SET_NULL, null=True, blank=True, related_name="sales")
    status = models.CharField("状態", max_length=12, choices=Status.choices, default=Status.PAID)
    subtotal = models.PositiveIntegerField("小計（税込円）")
    total = models.PositiveIntegerField("合計（税込円）")
    created_at = models.DateTimeField("販売日時", auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]
        verbose_name = "販売"
        verbose_name_plural = "販売"

    def __str__(self):
        return self.receipt_number or f"販売 #{self.pk}"


class SaleItem(models.Model):
    sale = models.ForeignKey(Sale, verbose_name="販売", on_delete=models.CASCADE, related_name="items")
    product = models.ForeignKey(Product, verbose_name="商品", on_delete=models.PROTECT, related_name="sale_items")
    product_name = models.CharField("商品名", max_length=120)
    unit_price = models.PositiveIntegerField("単価（税込円）")
    quantity = models.PositiveIntegerField("数量")
    line_total = models.PositiveIntegerField("金額（税込円）")

    class Meta:
        verbose_name = "販売明細"
        verbose_name_plural = "販売明細"


class Payment(models.Model):
    class Method(models.TextChoices):
        CASH = "cash", "現金"
        PAYPAY = "paypay", "PayPay"
        ALIPAY = "alipay", "支付宝 / Alipay"
        CARD = "card", "クレジットカード（デモ）"

    sale = models.OneToOneField(Sale, verbose_name="販売", on_delete=models.CASCADE, related_name="payment")
    method = models.CharField("支払方法", max_length=16, choices=Method.choices)
    amount = models.PositiveIntegerField("支払額（税込円）")
    transaction_reference = models.CharField("取引参照番号", max_length=40, unique=True)
    paid_at = models.DateTimeField("支払日時", auto_now_add=True)

    class Meta:
        ordering = ["-paid_at"]
        verbose_name = "決済"
        verbose_name_plural = "決済"


def create_demo_reference():
    return f"DEMO-{uuid.uuid4().hex[:16].upper()}"
