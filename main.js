//import * as poi from './build/Release/obj/poi-binding';
import java from 'java';
java.classpath.push('../poi-binding/src/poi-bin-5.2.0/poi-5.2.0.jar')
var myClass = java.newInstanceSync('api.model.MyObject');

console.log({ myClass })