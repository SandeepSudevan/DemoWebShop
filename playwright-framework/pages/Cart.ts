import { Page, Locator, Expect   } from '@playwright/test';

export class Cart {
    cartLink: Locator;
    productName: Locator;
    cartButton: Locator;
    termsOfService: Locator;
    constructor(private page: Page, private expect: Expect) {
        this.expect = expect;
        this.page = page;
        this.cartLink = page.locator('#topcartlink');
        this.productName = page.locator('td.product');
        this.termsOfService = page.locator('#termsofservice');
        this.cartButton = page.getByRole('button', { name: 'Checkout' });
    }
    async goToCart() {
        await this.cartLink.click();
    }

    async checkCart(product: string) {
        await this.expect(this.productName.filter({hasText: product})).toContainText(product);
    }

    async checkCartQuantity() {
        await this.termsOfService.check();
        await Promise.all([
            this.page.waitForURL(/checkout/i, { timeout: 15000 }),
            this.cartButton.click(),
        ]);
    }

}