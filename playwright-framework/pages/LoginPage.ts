import { Page, Locator } from '@playwright/test';

export class LoginPage {
    loginLink: Locator;
    username: Locator;
    password: Locator;
    loginButton: Locator;
    searchTextbox: Locator;
    searchButton: Locator;
    item: Locator;
    addtoCart: Locator;
    constructor(private page: Page) {
        this.page = page;
        this.loginLink = page.getByRole('link', { name: 'Log in' });
        this.username = page.getByRole('textbox', { name: 'Email' });
        this.password = page.getByRole('textbox', { name: 'Password' });
        this.loginButton = page.getByRole('button', { name: 'Log in' });
        this.searchTextbox = page.locator('#small-searchterms');
        this.searchButton = page.getByRole('button', { name: 'Search' });
        this.item = page.locator('div.product-item');
        this.addtoCart = page.locator('.add-to-cart').getByRole('button', { name: 'Add to cart' });
    }

    async login(username: string, password: string) {
        await this.loginLink.click();
        await this.username.fill(username);
        await this.password.fill(password);
        await this.loginButton.click();
    }

    async siteURL(url: string) {
        await this.page.goto(url);
    }
    async search(term: string, product: string) {
        await this.searchTextbox.fill(term);
        await this.searchButton.click();
        await this.item.filter({ hasText: product }).click();
        await this.addtoCart.click();
    }
}   