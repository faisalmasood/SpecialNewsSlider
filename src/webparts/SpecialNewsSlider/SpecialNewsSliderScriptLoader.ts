import { SPComponentLoader } from "@microsoft/sp-loader";
import JQueryStatic = require('jquery');
declare var window: any;
declare var jQuery: JQueryStatic;
declare var $: JQueryStatic;

/**
 * @class
 * This class is used to resolve the js script conflict issue.
 */
namespace SpecialNewsSliderScriptLoader {

    /**
     * @interface
     * The js script object interface
     */
    export interface IScript {
        /**
         * The Url of the script, which mustn't be local path.
         */
        Url: string;
        /**
        * The script's global exports name. 
        */
        GlobalExportsName: string;
        /**
        * After the js script loaded, which object have been extended property, function...
        * Such as after jquery script loaded, the jQuery will be attached to window object we can use window.jQuery to access jquery functions. 
        */
        ExtendedTarget: any;
        /**
         * The extended property or function chain.
         * If your script will attach property(X) to window.A.B.C.D
         * Here please set value as A.B.C.D
         */
        ExtendedPropertyOrFunctionChain: string;
    }

    /**
     * @function
     * The exported function, which is used to load the script and scripts who is depend on the script.
     * If purpose is loading jquery and jquery UI, the script will be jquery and denpendencies should contain jquery ui.
     * @param script The script that need to be firstly load.
     * @param dependencies The scripts that depend on the script.
     */
    export function LoadScript(script: IScript, dependencies: IScript[]): Promise<void> {
        let scriptObject = GetExtendedPropertyOrFunction(script.ExtendedTarget, script.ExtendedPropertyOrFunctionChain);
        if (scriptObject != undefined) {
            if (script.ExtendedPropertyOrFunctionChain.toLowerCase() == "jquery") {
                jQuery = window.jQuery as JQueryStatic as JQueryStatic;
            }
            return LoadDependencies(dependencies);
        }
        else {
            if (script.GlobalExportsName != null && script.GlobalExportsName != '') {
                return SPComponentLoader.loadScript(script.Url, { globalExportsName: script.GlobalExportsName }).then(() => {
                    return LoadDependencies(dependencies);
                });
            }
            else {
                return SPComponentLoader.loadScript(script.Url).then(() => {
                    return LoadDependencies(dependencies);
                });
            }
        }
    }

    /**
     * @function
     * The exported function, which is use to wait specific property or function have been attached to target object.
     * The timeout value is 10 seconds.
     * @param extendedTarget After the js script loaded, which object have been extended property, function...
     * @param extendedTargetPropertyOrFunctionChain The extended property or function chain.
     */
    export function WaitExtendedObjectAttached(extendedTarget: object, extendedTargetPropertyOrFunctionChain: string): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            var waitTimeout = new Date().getTime() + 10000;
            let waitingInternal = setInterval(function () {
                if (new Date().getTime() > waitTimeout) {
                    clearInterval(waitingInternal);
                    resolve(false);
                }
                else {
                    var propertyOrFunc = GetExtendedPropertyOrFunction(extendedTarget, extendedTargetPropertyOrFunctionChain);
                    if (propertyOrFunc) {
                        clearInterval(waitingInternal);
                        resolve(true);
                    }
                }
            }, 100);
        });
    }

    /**
     * @function
     * The private function that checks whether the scripts are loaded or not. Only load didn't load scripts by using SPComponentLoader.
     * Note: For such scripts, we didn't wait until the related properties/functions attached to target object, please use WaitExtendedObjectAttached method to wait for ready use.
     * @param dependencies The dependencies scripts that need to be loaded.
     */
    function LoadDependencies(dependencies: IScript[]): Promise<void> {
        dependencies.forEach(script => {
            let scriptObject = GetExtendedPropertyOrFunction(script.ExtendedTarget, script.ExtendedPropertyOrFunctionChain);
            if (scriptObject == undefined) {
                if (script.GlobalExportsName != null && script.GlobalExportsName != '') {
                    SPComponentLoader.loadScript(script.Url, { globalExportsName: script.GlobalExportsName });
                }
                else {
                    SPComponentLoader.loadScript(script.Url);
                }
            }
        });
        return Promise.resolve();
    }

    /**
     * @function
     * The private function that used to get target attached property value, if the property or function was not exists, return undefined.
     * @param extendedTarget The extended target object such as window
     * @param ExtendedPropertyOrFunctionChain The extended property or function chain such as jQuery, if the property or function is nested one, please using format A.B.C
     */
    function GetExtendedPropertyOrFunction(extendedTarget: object, ExtendedPropertyOrFunctionChain: string): any {
        let object = undefined;
        if (ExtendedPropertyOrFunctionChain != null && ExtendedPropertyOrFunctionChain != '') {
            var chains = ExtendedPropertyOrFunctionChain.split('.');
            for (var i = 0; i < chains.length; i++) {
                object = GetPropertyOrFunction(object != undefined ? object : extendedTarget, chains[i]);
                if (object == undefined) {
                    break;
                }
            }
        }
        return object;
    }

    /**
     * @function
     * The private function that used to get property value or function in target object's top level.
     * @param target The target object used to retrieve property or function.
     * @param propertyOrFunctionName The property name or function name.
     */
    function GetPropertyOrFunction(target: any, propertyOrFunctionName: string): any {
        let propertyOrFunction = undefined;
        if (target != undefined) {
            for (var p in target) {
                if (p.toLocaleLowerCase() == propertyOrFunctionName.toLocaleLowerCase()) {
                    propertyOrFunction = target[p];
                    break;
                }
            }
        }
        return propertyOrFunction;
    }
}

export default SpecialNewsSliderScriptLoader;