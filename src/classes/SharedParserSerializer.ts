import {JSDOM} from 'jsdom';

export default class jsDOM {

    private static dom = new JSDOM();
    public static parser = new this.dom.window.DOMParser();
    public static serializer = new this.dom.window.XMLSerializer();
}