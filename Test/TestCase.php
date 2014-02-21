<?php
/**
 * @author		  Ronald Marske <r.marske@secu-ring.de>
 * @filesource	  Test/TestCase.php
 *
 * @copyright     Copyright (c) 2013 Ronald Marske, All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions
 * are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in
 *       the documentation and/or other materials provided with the
 *       distribution.
 *
 *     * Neither the name of Ronald Marske nor the names of his
 *       contributors may be used to endorse or promote products derived
 *       from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 * "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 * LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 * FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 * COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 * INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
 * CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
 * LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
 * ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */

namespace Test;

use Excellence\Workbook;
use Test\Excellence\Stub\DataSource;
use Excellence\Delegates\WorkbookDelegate;
use Test\Excellence\Stub\MergeableDataSource;
use Test\Excellence\Stub\StylableDataSource;

class TestCase extends \PHPUnit_Framework_TestCase {

#pragma mark - creations

	/**
	 * @param string $sIdentifier
	 * @param WorkbookDelegate $oDelegate
	 *
	 * @return Workbook
	 */
	public function makeWorkbook($sIdentifier = 'workbook', WorkbookDelegate $oDelegate = null) {

		if (null == $oDelegate) {
			$oDelegate  = $this->makeDelegate();
		}

		return new Workbook($sIdentifier, $oDelegate);
	}

	/**
	 * @return DataSource
	 */
	public function makeDelegate() {
		return new DataSource();
	}

	/**
	 * @return MergableDataSource
	 */
	public function makeMergeDelegate() {
		return new MergeableDataSource();
	}

	/**
	 * @return StylableDataSource
	 */
	public function makeStylableDelegate() {
		return new StylableDataSource();
	}

#pragma mark - dataProvider

	/**
	 * data provider that returns values that don't match or could
	 * type casted to positive integer values.
	 *
	 * @return array
	 */
	public function dataProviderInvalidValuesForPositiveIntegers() {
		return array(
			array(0),
			array(-2),
			array('0'),
			array('0.9'),
			array('-1'),
			array('test'),
			array(array()),
			array(false),
		);
	}
} 