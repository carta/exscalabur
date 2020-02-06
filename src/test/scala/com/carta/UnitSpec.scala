package com.carta

import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers
import org.scalatest.{Inside, Inspectors, OptionValues}

abstract class UnitSpec extends AnyFlatSpec with Matchers with OptionValues with Inside with Inspectors