package test

import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.{Inside, Inspectors, OptionValues}
import org.scalatest.matchers.must.Matchers

abstract class UnitSpec extends AnyFlatSpec with Matchers with OptionValues with Inside with Inspectors