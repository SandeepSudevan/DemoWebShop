import  { Expect, Page } from '@playwright/test';
import  { LoginPage } from '../pages/LoginPage';
import  { Cart } from '../pages/Cart';
import  { Checkout } from '../pages/Checkout';
export class POManager {
    loginPage: LoginPage;
    cart: Cart;
    checkout: Checkout;
    constructor(private page: Page, private expect: Expect) {
        this.loginPage = new LoginPage(page);
        this.cart = new Cart(page, expect);
        this.checkout = new Checkout(page, expect);
    }
}