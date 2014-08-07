/**
 * Created by bohnen on 2014/08/07.
 */
lst = [1,2]
lst.with {
    add(3)
    add(4)
    println size()
    println contains(2)
}